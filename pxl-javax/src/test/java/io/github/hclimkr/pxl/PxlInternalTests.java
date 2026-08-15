package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.constraint.PxlByteSize;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.codec.PxlCellResolver;
import io.github.hclimkr.pxl.internal.constraint.PxlByteSizeValidator;
import io.github.hclimkr.pxl.internal.core.PxlContentsHandler;
import io.github.hclimkr.pxl.internal.core.PxlCoreCsvExporter;
import io.github.hclimkr.pxl.internal.core.PxlCoreExcelExporter;
import io.github.hclimkr.pxl.internal.i18n.PxlI18n;
import io.github.hclimkr.pxl.internal.meta.*;
import io.github.hclimkr.pxl.internal.support.*;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import io.github.hclimkr.pxl.type.PxlOptionalBoolean;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.StringWriter;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.Period;
import java.util.*;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Tests that call {@code internal/} members directly rather than through the public API.
 * <p>
 * These are not consumer-facing entry points, so they belong neither with the public {@code util/} tests
 * ({@link PxlUtilityTests}) nor with a feature suite: they pin behavior the public API depends on but does not
 * expose, which is why reaching it needs an {@code internal.*} import. Everything covered here may change
 * without notice along with the internals themselves.
 * <p>
 * Covered areas: {@code internal/codec} (the export dispatcher's string form), {@code internal/support}
 * (reflection, class, number, date-time, temporal-amount, workbook, type-reference helpers),
 * {@code internal/meta} (converter metadata and the workbook/sheet/column factory chain), {@code internal/core}
 * (contents handler, exporter argument validation), {@code internal/i18n} and {@code internal/constraint}.
 */
public class PxlInternalTests {

    // ==================================================================
    // internal/codec - PxlCellResolver: the export dispatcher's string form
    // ==================================================================

    @Test
    public void buildDataString_everyColumnType_rendersWithoutACell() throws Exception {
        // Every codec must produce its string with cell == null, including the seven primitive codecs that
        // PxlCollectionCodec never reaches (a Collection cannot hold a primitive), so none of them may
        // dereference the cell before the null guard.
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMeta(AllTypesWorkbook.class, null);

        try {
            final List<PxlExportSheetMeta> sheetMetas =
                    PxlExportSheetMeta.makeExportSheetMetas(AllTypesWorkbook.class, workbookMeta, null, true);
            final List<PxlExportColumnMeta> columnMetas =
                    PxlExportColumnMeta.makeExportColumnMetas(sheetMetas.get(0), true);
            assertThat(columnMetas).hasSizeGreaterThan(30);   // all-types fixture: every supported column type

            for (final PxlExportColumnMeta columnMeta : columnMetas) {
                // exportSample is the string an export receives for this column, so it exercises the real path.
                final String rendered = PxlCellResolver.buildDataString(columnMeta.getExportSample(), columnMeta);

                assertThat(rendered).as(columnMeta.getActualExportColumnName()).isNotNull();
            }
        } finally {
            PxlWorkbookUtils.closeWorkbook(workbookMeta.getWorkbook());
        }
    }

    @Test
    public void buildDataCell_withAndWithoutCell_agreeOnTheString() throws Exception {
        // The cell-based and cell-less calls must not drift apart: what is written into the cell is what the
        // string form returns.
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMeta(AllTypesWorkbook.class, null);

        try {
            final List<PxlExportSheetMeta> sheetMetas =
                    PxlExportSheetMeta.makeExportSheetMetas(AllTypesWorkbook.class, workbookMeta, null, true);
            final List<PxlExportColumnMeta> columnMetas =
                    PxlExportColumnMeta.makeExportColumnMetas(sheetMetas.get(0), true);

            final Sheet sheet = workbookMeta.getWorkbook().createSheet("Agree");
            final Row row = sheet.createRow(0);

            int columnIndex = 0;
            for (final PxlExportColumnMeta columnMeta : columnMetas) {
                final Cell cell = row.createCell(columnIndex++);
                final String written = PxlCellResolver.buildDataCell(cell, columnMeta.getExportSample(), columnMeta);
                final String rendered = PxlCellResolver.buildDataString(columnMeta.getExportSample(), columnMeta);

                assertThat(rendered).as(columnMeta.getActualExportColumnName()).isEqualTo(written);
            }
        } finally {
            PxlWorkbookUtils.closeWorkbook(workbookMeta.getWorkbook());
        }
    }

    @Test
    public void buildDataString_nullValueAndNullMeta_followTheNullContract() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMeta(AllTypesWorkbook.class, null);

        try {
            final List<PxlExportSheetMeta> sheetMetas =
                    PxlExportSheetMeta.makeExportSheetMetas(AllTypesWorkbook.class, workbookMeta, null, true);
            final PxlExportColumnMeta columnMeta =
                    PxlExportColumnMeta.makeExportColumnMetas(sheetMetas.get(0), true).get(0);

            // A null value, and a blank string, both yield the column's export-null string rather than null:
            // a cell-less format needs a field to write, not a missing one.
            assertThat(PxlCellResolver.buildDataString(null, columnMeta)).isEqualTo(columnMeta.getExportNullString());
            assertThat(PxlCellResolver.buildDataString("   ", columnMeta)).isEqualTo(columnMeta.getExportNullString());

            // Without metadata there is nothing to render by.
            assertThat(PxlCellResolver.buildDataString("anything", null)).isNull();
        } finally {
            PxlWorkbookUtils.closeWorkbook(workbookMeta.getWorkbook());
        }
    }


    // ==================================================================
    // internal/support helpers (reflection, class, number, date-time, temporal-amount, workbook, type-reference)
    // ==================================================================

    // ------------------------------------------------------------------
    // PxlReflectionSupport
    // ------------------------------------------------------------------

    @Test
    public void reflectionSupport_withAnnotationValue_overridesSingleMember() throws Exception {
        final PxlColumn original = AllTypesRow.class.getDeclaredField("text").getAnnotation(PxlColumn.class);
        assertThat(original.name()).isEqualTo(new String[]{"Text"});

        final PxlColumn overridden = PxlReflectionSupport.withAnnotationValue(original, "name", new String[]{"Renamed"});
        assertThat(overridden.name()).isEqualTo(new String[]{"Renamed"});           // overridden member
        assertThat(overridden.exportSample()).isEqualTo(original.exportSample());   // other members delegate to the original
        assertThat(overridden.annotationType()).isEqualTo(PxlColumn.class);         // annotationType() delegates too
    }

    @Test
    public void reflectionSupport_withAnnotationValue_invalidKeyOrValue_throws() throws Exception {
        final PxlColumn annotation = AllTypesRow.class.getDeclaredField("text").getAnnotation(PxlColumn.class);
        // no such member
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.withAnnotationValue(annotation, "noSuchMember", "x"));
        // null new value
        assertThrows(PxlArgumentException.class, () -> PxlReflectionSupport.withAnnotationValue(annotation, "name", null));
        // incompatible value type (name() returns String, given an Integer)
        assertThrows(PxlArgumentException.class, () -> PxlReflectionSupport.withAnnotationValue(annotation, "name", 123));
    }

    @Test
    public void reflectionSupport_stringTypeConstructor_detectedAndReturned() {
        assertThat(PxlReflectionSupport.hasStringTypeConstructor(Point.class)).isTrue();   // Point(String) exists
        assertThat(PxlReflectionSupport.hasStringTypeConstructor(Object.class)).isFalse();
        assertThat(PxlReflectionSupport.getStringTypeConstructor(Point.class)).isNotNull();
        assertThat(PxlReflectionSupport.getStringTypeConstructor(Object.class)).isNull();
    }

    @Test
    public void reflectionSupport_newClassInstance_errorPaths_throw() throws Exception {
        // no no-arg constructor -> ReflectiveOperationException path
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.newClassInstance(NoDefaultCtorRow.class));
        // constructor itself throws -> InvocationTargetException path (cause propagated)
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.newClassInstance(ThrowingCtorRow.class));
        // happy path
        assertThat(PxlReflectionSupport.newClassInstance(Point.class)).isInstanceOf(Point.class);
    }

    @Test
    public void reflectionSupport_getParameterizedArgument0_variants() throws Exception {
        final Field stringList = AllTypesRow.class.getDeclaredField("stringList");   // List<String>
        assertThat(PxlReflectionSupport.getParameterizedArgument0(stringList)).isEqualTo(String.class);

        // a non-parameterized (raw) field is rejected
        final Field text = AllTypesRow.class.getDeclaredField("text");   // String
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.getParameterizedArgument0(text));

        // a nested generic (List<List<String>>) whose first argument is not a concrete Class is rejected
        final Field nested = NestedCollectionRow.class.getDeclaredField("nested");
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.getParameterizedArgument0(nested));
    }

    @Test
    public void reflectionSupport_fieldValue_getSet_andErrors() throws Exception {
        final Field textField = AllTypesRow.class.getDeclaredField("text");
        final AllTypesRow row = new AllTypesRow();

        PxlReflectionSupport.setFieldValue(textField, row, "hello");
        assertThat(PxlReflectionSupport.getFieldValue(textField, row)).isEqualTo("hello");

        // incompatible value type -> wrapped as PxlReflectionException
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.setFieldValue(textField, row, 123));
        // wrong target object type -> wrapped as PxlReflectionException
        assertThrows(PxlReflectionException.class, () -> PxlReflectionSupport.getFieldValue(textField, "not a row"));
    }

    @Test
    public void reflectionSupport_getToStringMethod_findsOverride() {
        assertThat(PxlReflectionSupport.getToStringMethod(Point.class)).isNotNull();   // Point overrides toString
        assertThat(PxlReflectionSupport.getToStringMethod(Object.class)).isNull();     // only Object's default
    }

    // ------------------------------------------------------------------
    // PxlClassSupport (concrete collection resolution)
    // ------------------------------------------------------------------

    @Test
    public void classSupport_getConcreteCollectionClass_resolvesInterfacesAndConcrete() throws Exception {
        assertThat(PxlClassSupport.getConcreteCollectionClass(List.class)).isEqualTo(ArrayList.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Collection.class)).isEqualTo(ArrayList.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Set.class)).isEqualTo(HashSet.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(SortedSet.class)).isEqualTo(TreeSet.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(NavigableSet.class)).isEqualTo(TreeSet.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Deque.class)).isEqualTo(ArrayDeque.class);
        assertThat(PxlClassSupport.getConcreteCollectionClass(Queue.class)).isEqualTo(LinkedList.class);
        // a concrete Collection class is returned as-is
        assertThat(PxlClassSupport.getConcreteCollectionClass(LinkedList.class)).isEqualTo(LinkedList.class);
    }

    @Test
    public void classSupport_getConcreteCollectionClass_unsupported_throws() {
        // a non-collection interface
        assertThrows(PxlReflectionException.class, () -> PxlClassSupport.getConcreteCollectionClass(Comparable.class));
        // a concrete non-collection class
        assertThrows(PxlReflectionException.class, () -> PxlClassSupport.getConcreteCollectionClass(String.class));
    }

    // ------------------------------------------------------------------
    // PxlNumberSupport (range / finiteness guards)
    // ------------------------------------------------------------------

    @Test
    public void numberSupport_requireWithinRange_variants() throws Exception {
        // BigDecimal overload
        assertThat(PxlNumberSupport.requireWithinRange(new BigDecimal("5"), 0, 10, "X")).isEqualByComparingTo("5");
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(new BigDecimal("-1"), 0, 10, "X"));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(new BigDecimal("11"), 0, 10, "X"));
        // Number overload
        assertThat(PxlNumberSupport.requireWithinRange((Number) 7, 0, 10, "X")).isEqualByComparingTo("7");
        // double overload
        assertThat(PxlNumberSupport.requireWithinRange(3.0, 0, 10, "X")).isEqualByComparingTo("3");
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(100.0, 0, 10, "X"));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireWithinRange(Double.NaN, 0, 10, "X"));
    }

    @Test
    public void numberSupport_parseFullyAsNumber_rejectsPartialConsumption() throws Exception {
        final DecimalFormat formatter = PxlNumberSupport.getDecimalFormat("#,##0");

        // what the pattern reads end to end
        assertThat(PxlNumberSupport.parseFullyAsNumber(formatter, "1,234", "X").intValue()).isEqualTo(1234);
        // read in part: DecimalFormat.parse(String) would return 123 / 1 and drop the rest without complaint
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.parseFullyAsNumber(formatter, "123abc", "X"));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.parseFullyAsNumber(formatter, "1e3", "X"));
        // read not at all: rejected before and after
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.parseFullyAsNumber(formatter, "abc", "X"));
    }

    @Test
    public void numberSupport_parseFullyAsBigDecimal_rejectsNonFiniteToken() throws Exception {
        final DecimalFormat formatter = PxlNumberSupport.getDecimalFormat("#,##0.00");
        formatter.setParseBigDecimal(true);   // what the meta does for BigDecimal/BigInteger columns

        assertThat(PxlNumberSupport.parseFullyAsBigDecimal(formatter, "1,234.50", "X")).isEqualByComparingTo("1234.50");
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.parseFullyAsBigDecimal(formatter, "1,234.50junk", "X"));
        // setParseBigDecimal(true) still yields a Double for the infinity and NaN tokens, which a bare cast to
        // BigDecimal would meet with a ClassCastException
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.parseFullyAsBigDecimal(formatter, "∞", "X"));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.parseFullyAsBigDecimal(formatter, "NaN", "X"));
    }

    @Test
    public void numberSupport_requireFiniteForExport_rejectsNonFinite() throws Exception {
        // finite and null values pass
        PxlNumberSupport.requireFiniteForExport((Float) 1.5F);
        PxlNumberSupport.requireFiniteForExport((Float) null);
        PxlNumberSupport.requireFiniteForExport((Double) 2.5);
        PxlNumberSupport.requireFiniteForExport((Double) null);
        // NaN / Infinity are rejected
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Float.NaN));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Float.POSITIVE_INFINITY));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Double.NaN));
        assertThrows(PxlCellCodecException.class, () -> PxlNumberSupport.requireFiniteForExport(Double.NEGATIVE_INFINITY));
    }

    // ------------------------------------------------------------------
    // PxlDateTimeSupport
    // ------------------------------------------------------------------

    @Test
    public void dateTimeSupport_javaDateLocalDateTime_roundTripAndNull() {
        assertThat(PxlDateTimeSupport.localDateTimeToJavaDate(null)).isNull();
        assertThat(PxlDateTimeSupport.javaDateToLocalDateTime(null)).isNull();

        final LocalDateTime ldt = LocalDateTime.of(2023, 6, 15, 10, 30, 45);
        final Date date = PxlDateTimeSupport.localDateTimeToJavaDate(ldt);
        assertThat(PxlDateTimeSupport.javaDateToLocalDateTime(date)).isEqualTo(ldt);
    }

    @Test
    public void dateTimeSupport_parseFullyAsDate_rejectsPartialConsumption() throws Exception {
        final SimpleDateFormat formatter = PxlDateTimeSupport.getCellSimpleDateFormatter("yyyy-MM-dd", Locale.ROOT);

        final Calendar calendar = Calendar.getInstance();
        calendar.setTime(PxlDateTimeSupport.parseFullyAsDate(formatter, "2024-01-02", "X"));
        assertThat(calendar.get(Calendar.YEAR)).isEqualTo(2024);
        assertThat(calendar.get(Calendar.DAY_OF_MONTH)).isEqualTo(2);

        // SimpleDateFormat.parse(String) would stop at the space and report 2 January 2024 as a success
        assertThrows(PxlCellCodecException.class, () -> PxlDateTimeSupport.parseFullyAsDate(formatter, "2024-01-02 xxx", "X"));
        assertThrows(PxlCellCodecException.class, () -> PxlDateTimeSupport.parseFullyAsDate(formatter, "not a date", "X"));
    }

    @Test
    public void dateTimeSupport_toExcelDateFormat_lowercasesOutsideLiterals() {
        assertThat(PxlDateTimeSupport.toExcelDateFormat("yyyy-MM-dd HH:mm:ss")).isEqualTo("yyyy-mm-dd hh:mm:ss");
        // text inside single quotes is copied verbatim (not lower-cased)
        assertThat(PxlDateTimeSupport.toExcelDateFormat("yyyy'T'HH")).isEqualTo("yyyy'T'hh");
    }

    // ------------------------------------------------------------------
    // PxlTemporalAmountSupport (DurationFormatUtils-style pattern parser)
    // ------------------------------------------------------------------

    @Test
    public void temporalAmountSupport_compileAndParse_durationAndPeriod() {
        final PxlTemporalAmountSupport.CompiledTemporalPattern durationPattern =
                PxlTemporalAmountSupport.compileTemporalPattern("HH:mm:ss");
        assertThat(PxlTemporalAmountSupport.parseDurationByPattern("01:02:03", durationPattern))
                .isEqualTo(Duration.ofHours(1).plusMinutes(2).plusSeconds(3));

        final PxlTemporalAmountSupport.CompiledTemporalPattern periodPattern =
                PxlTemporalAmountSupport.compileTemporalPattern("y'y'M'm'd'd'");
        assertThat(PxlTemporalAmountSupport.parsePeriodByPattern("1y2m3d", periodPattern))
                .isEqualTo(Period.of(1, 2, 3));
    }

    @Test
    public void temporalAmountSupport_compile_literalsAndEscapedQuote() {
        // "d''H" is: token d, an escaped literal single quote (''), token H.
        final PxlTemporalAmountSupport.CompiledTemporalPattern pattern =
                PxlTemporalAmountSupport.compileTemporalPattern("d''H");
        assertThat(PxlTemporalAmountSupport.parseDurationByPattern("5'7", pattern))
                .isEqualTo(Duration.ofDays(5).plusHours(7));
    }

    @Test
    public void temporalAmountSupport_parse_mismatchAndOverflow_throw() {
        final PxlTemporalAmountSupport.CompiledTemporalPattern hms =
                PxlTemporalAmountSupport.compileTemporalPattern("HH:mm:ss");
        // the value does not match the pattern
        assertThrows(IllegalArgumentException.class, () -> PxlTemporalAmountSupport.parseDurationByPattern("nope", hms));

        final PxlTemporalAmountSupport.CompiledTemporalPattern days = PxlTemporalAmountSupport.compileTemporalPattern("d");
        // a Period field overflows the int range (Math.toIntExact)
        assertThrows(IllegalArgumentException.class, () -> PxlTemporalAmountSupport.parsePeriodByPattern("99999999999", days));
        // a Duration overflows its range
        assertThrows(IllegalArgumentException.class,
                () -> PxlTemporalAmountSupport.parseDurationByPattern(String.valueOf(Long.MAX_VALUE), days));
    }

    // ------------------------------------------------------------------
    // PxlWorkbookSupport (public helpers)
    // ------------------------------------------------------------------

    @Test
    public void workbookSupport_makeUniqueSafeSheetName_avoidsCollision() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Data");
            assertThat(PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, "Fresh")).isEqualTo("Fresh");
            // a colliding name gets a " (2)" suffix (POI matches case-insensitively)
            assertThat(PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, "Data")).isEqualTo("Data (2)");
            workbook.createSheet("Data (2)");
            assertThat(PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, "Data")).isEqualTo("Data (3)");
        }
    }

    @Test
    public void workbookSupport_makeUniqueDefinedName_avoidsCollision() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            assertThat(PxlWorkbookSupport.makeUniqueDefinedName(workbook, "MyName")).isEqualTo("MyName");

            final Name name = workbook.createName();
            name.setNameName("MyName");
            name.setRefersToFormula("Sheet1!$A$1");

            assertThat(PxlWorkbookSupport.makeUniqueDefinedName(workbook, "MyName")).isEqualTo("MyName_2");
        }
    }

    @Test
    public void workbookSupport_createWorkbook_perEngine() throws Exception {
        assertThat(PxlWorkbookSupport.createWorkbook(PxlExcelEngine.XSSF, 100)).isInstanceOf(XSSFWorkbook.class);
        assertThat(PxlWorkbookSupport.createWorkbook(PxlExcelEngine.HSSF, 100)).isInstanceOf(HSSFWorkbook.class);
        assertThat(PxlWorkbookSupport.createWorkbook(PxlExcelEngine.SXSSF, 100)).isInstanceOf(SXSSFWorkbook.class);
    }

    @Test
    public void workbookSupport_validateWorkbookNameFieldType_rejectsNonString() throws Exception {
        // null and a String-typed name field are accepted
        PxlWorkbookSupport.validateWorkbookNameFieldType(null);
        PxlWorkbookSupport.validateWorkbookNameFieldType(AllTypesWorkbook.class);   // String workbookName -> ok
        // a non-String @PxlWorkbookName field is rejected
        assertThrows(PxlDataException.class, () -> PxlWorkbookSupport.validateWorkbookNameFieldType(BadWorkbookNameWorkbook.class));
    }

    // ------------------------------------------------------------------
    // PxlTypeReference (super type token)
    // ------------------------------------------------------------------

    @Test
    @SuppressWarnings("deprecation")
    public void typeReference_capturesGenericArgument() {
        final PxlTypeReference<List<String>> reference = new PxlTypeReference<List<String>>() {
        };
        assertThat(reference.getType()).isInstanceOf(ParameterizedType.class);
        assertThat(((ParameterizedType) reference.getType()).getRawType()).isEqualTo(List.class);
    }

    @Test
    @SuppressWarnings({"deprecation", "rawtypes"})
    public void typeReference_rawInstantiation_throws() {
        // a raw (non-parameterized) subclass has no captured type argument -> rejected
        assertThrows(IllegalArgumentException.class, () -> new PxlTypeReference() {
        });
    }

    // ==================================================================
    // internal/meta helpers (custom-converter resolution, factory chain, toString)
    // ==================================================================

    @Test
    public void exportConverterMeta_of_resolvesAndValidates() throws Exception {
        // instance converter (Money) and toString-based (Point) both resolve
        assertThat(PxlExportColumnMeta.PxlExportConverterMeta.of(Money.class)).isNotNull();
        assertThat(PxlExportColumnMeta.PxlExportConverterMeta.of(Point.class)).isNotNull();
        // neither a converter nor a toString override -> unsupported
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(Object.class));
    }

    @Test
    public void exportConverterMeta_of_invalidSignatures_throw() {
        // return type is not String
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(BadExportConverterReturnObject.class));
        // static converter with 0 args (must take exactly the object)
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(BadExportConverterObject.class));
        // instance converter that takes an argument (must take none)
        assertThrows(PxlArgumentException.class,
                () -> PxlExportColumnMeta.PxlExportConverterMeta.of(BadExportConverterInstanceParamObject.class));
    }

    @Test
    public void importConverterMeta_of_resolvesAndValidates() throws Exception {
        // static converter (Money) and String-constructor-based (Point) both resolve
        assertThat(PxlImportColumnMeta.PxlImportConverterMeta.of(Money.class)).isNotNull();
        assertThat(PxlImportColumnMeta.PxlImportConverterMeta.of(Point.class)).isNotNull();
        // an enum resolves even without a converter or a String constructor
        assertThat(PxlImportColumnMeta.PxlImportConverterMeta.of(Grade.class)).isNotNull();
        // neither a converter, a String constructor, nor an enum -> unsupported
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(Object.class));
    }

    @Test
    public void importConverterMeta_of_invalidSignatures_throw() {
        // return type is not the object type
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(BadImportConverterReturnObject.class));
        // converter is an instance method (must be static)
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(BadImportConverterInstanceObject.class));
        // converter does not take a single String argument
        assertThrows(PxlArgumentException.class,
                () -> PxlImportColumnMeta.PxlImportConverterMeta.of(BadImportConverterParamObject.class));
    }

    @Test
    public void exportMeta_factoryChain_andToString() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMeta(AllTypesWorkbook.class, null);
        final List<PxlExportSheetMeta> sheetMetas =
                PxlExportSheetMeta.makeExportSheetMetas(AllTypesWorkbook.class, workbookMeta, null, false);
        assertThat(sheetMetas).isNotEmpty();

        final PxlExportSheetMeta sheetMeta = sheetMetas.get(0);
        assertThat(sheetMeta.toString()).isNotBlank();   // PxlExportSheetMeta.toString

        final List<PxlExportColumnMeta> columnMetas = PxlExportColumnMeta.makeExportColumnMetas(sheetMeta, false);
        assertThat(columnMetas).isNotEmpty();
        assertThat(columnMetas.get(0).toString()).isNotBlank();   // PxlExportColumnMeta.toString
    }

    @Test
    public void importMeta_factoryChain_andToString() throws Exception {
        final PxlImportWorkbookMeta workbookMeta =
                PxlImportWorkbookMeta.makeImportWorkbookMeta(AllTypesWorkbook.class, null);
        // instance sheet-option accessor: an empty option list yields null at any index
        assertThat(workbookMeta.getImportSheetOption(0)).isNull();

        final List<PxlImportSheetMeta> sheetMetas =
                PxlImportSheetMeta.makeImportSheetMetas(AllTypesWorkbook.class, workbookMeta, null);
        assertThat(sheetMetas).isNotEmpty();

        final PxlImportSheetMeta sheetMeta = sheetMetas.get(0);
        assertThat(sheetMeta.toString()).isNotBlank();   // PxlImportSheetMeta.toString

        final List<PxlImportColumnMeta> columnMetas = PxlImportColumnMeta.makeImportColumnMetas(sheetMeta);
        assertThat(columnMetas).isNotEmpty();
        assertThat(columnMetas.get(0).toString()).isNotBlank();   // PxlImportColumnMeta.toString
    }

    // ==================================================================
    // internal/core helpers
    // ==================================================================

    @Test
    @SuppressWarnings({"deprecation", "unchecked"})
    public void contentsHandler_accumulatesHeaderAndPaddedDataRows() throws Exception {
        // PxlContentsHandler is a POI event-model SheetContentsHandler (deprecated/reference-only).
        // Drive the SAX-style callbacks directly: row 0 is the header, later rows are padded to header width,
        // and skipped cells within a row are padded with empty strings.
        final PxlContentsHandler handler = new PxlContentsHandler();

        handler.startRow(0);
        handler.cell("A1", "H1", null);
        handler.cell("C1", "H3", null);   // skips B1 -> padded
        handler.endRow(0);

        handler.startRow(1);
        handler.cell("A2", "v1", null);   // shorter than the header -> padded on endRow
        handler.endRow(1);

        handler.headerFooter("ignored", true, "oddHeader");   // no-op
        handler.endSheet();                                    // no-op

        // The class exposes no getters, so read the accumulated state reflectively.
        final List<String> header = (List<String>) PxlReflectionSupport.getFieldValue(
                PxlContentsHandler.class.getDeclaredField("header"), handler);
        assertThat(header).containsExactly("H1", "", "H3");

        final List<List<String>> rows = (List<List<String>>) PxlReflectionSupport.getFieldValue(
                PxlContentsHandler.class.getDeclaredField("rows"), handler);
        assertThat(rows).hasSize(1);
        assertThat(rows.get(0)).containsExactly("v1", "", "");   // padded to the header width (3)
    }

    @Test
    public void excelExporter_singleSheetBuildWorkbook_writesSheet() throws Exception {
        // The single-sheet String overload of buildWorkbook is not used by the fluent builder (which always
        // uses the multi-sheet list overload), so exercise it directly.
        final PxlExportWorkbookMeta workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null);
        final List<Employee> rows = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, null, Grade.B, "Sales"));

        final Workbook workbook = PxlCoreExcelExporter.buildWorkbook("People", rows, Employee.class, workbookMeta, null);

        final Sheet sheet = workbook.getSheet("People");
        assertThat(sheet).isNotNull();
        assertThat(sheet.getRow(0)).as("header row").isNotNull();
        assertThat(sheet.getLastRowNum()).isEqualTo(2);   // header + 2 data rows (0-based last index 2)
    }

    @Test
    public void excelExporter_multiSheetBuildWorkbook_validatesArguments() throws Exception {
        final List<Employee> employees = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"));
        final List<Collection<?>> objects = Arrays.<Collection<?>>asList(employees);
        final List<Class<?>> classes = Arrays.<Class<?>>asList(Employee.class);

        // sheetNames vs sheetObjects count mismatch
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1", "S2"), objects, classes, PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
        // sheetNames vs rowClasses count mismatch
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1"), objects, Arrays.<Class<?>>asList(Employee.class, Employee.class),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
        // a null sheet-object element
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1"), Arrays.<Collection<?>>asList((Collection<?>) null), classes,
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
        // a null row-class element
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildWorkbook(
                Arrays.asList("S1"), objects, Arrays.<Class<?>>asList((Class<?>) null),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null), null));
    }

    @Test
    public void excelExporter_multiSheetBuildSampleWorkbook_validatesArguments() throws Exception {
        final List<Class<?>> classes = Arrays.<Class<?>>asList(Employee.class);

        // sheetNames vs rowClasses count mismatch
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildSampleWorkbook(
                Arrays.asList("S1", "S2"), classes, PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null)));
        // duplicate sheet name
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildSampleWorkbook(
                Arrays.asList("Dup", "Dup"), Arrays.<Class<?>>asList(Employee.class, Employee.class),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null)));
        // a null row-class element
        assertThrows(PxlDataException.class, () -> PxlCoreExcelExporter.buildSampleWorkbook(
                Arrays.asList("S1"), Arrays.<Class<?>>asList((Class<?>) null),
                PxlExportWorkbookMeta.makeExportWorkbookMeta(null, null)));
    }

    // ==================================================================
    // internal/i18n (shared UTF-8 resource-bundle loader)
    // ==================================================================

    @Test
    public void i18n_getBundleAndMessage_resolveDisableAndFallback() throws Exception {
        // i18n disabled (blank base name or null locale) -> null bundle
        assertThat(PxlI18n.getBundle("", Locale.ENGLISH)).isNull();
        assertThat(PxlI18n.getBundle("messages", (Locale) null)).isNull();
        assertThat(PxlI18n.getBundle("")).isNull();

        // a present bundle resolves; a known key returns its (locale-specific) message
        assertThat(PxlI18n.getBundle("messages", Locale.ENGLISH)).isNotNull();
        assertThat(PxlI18n.getMessage("messages", Locale.ENGLISH, "staff.column.role")).isEqualTo("Role");
        assertThat(PxlI18n.getMessage("messages", Locale.KOREAN, "staff.column.role")).isEqualTo("역할");
        // a missing key falls back to the key itself
        assertThat(PxlI18n.getMessage("messages", Locale.ENGLISH, "no.such.key")).isEqualTo("no.such.key");
        // i18n disabled -> the key is returned unchanged
        assertThat(PxlI18n.getMessage("", null, "staff.column.role")).isEqualTo("staff.column.role");
        // params overload (the message has no placeholder, so it is returned as-is)
        assertThat(PxlI18n.getMessage("messages", Locale.ENGLISH, "staff.sheet", new Object[]{"x"})).isEqualTo("Staff");

        // JVM-default-locale overload resolves the bundle
        assertThat(PxlI18n.getBundle("messages")).isNotNull();

        // a missing bundle -> PxlI18nException
        assertThrows(PxlI18nException.class, () -> PxlI18n.getBundle("no-such-bundle", Locale.ENGLISH));
    }

    // ==================================================================
    // internal/constraint (PxlByteSize validator initialization)
    // ==================================================================

    @Test
    public void byteSizeValidator_initialize_rejectsInvalidBounds() throws Exception {
        final PxlByteSize annotation = ByteSizeRow.class.getDeclaredField("code").getAnnotation(PxlByteSize.class);
        final PxlByteSizeValidator validator = new PxlByteSizeValidator();

        // a valid annotation (min=0, max=5) initializes without error
        validator.initialize(annotation);

        // min < 0
        assertThrows(IllegalArgumentException.class,
                () -> validator.initialize(PxlReflectionSupport.withAnnotationValue(annotation, "min", -1)));
        // max < 0
        assertThrows(IllegalArgumentException.class,
                () -> validator.initialize(PxlReflectionSupport.withAnnotationValue(annotation, "max", -1)));
        // max < min
        assertThrows(IllegalArgumentException.class,
                () -> validator.initialize(PxlReflectionSupport.withAnnotationValue(
                        PxlReflectionSupport.withAnnotationValue(annotation, "min", 10), "max", 3)));
    }

    // ------------------------------------------------------------------
    // internal/meta - CSV export metadata
    //
    // CSV export has the sheet form only, which builds its metadata with no workbook class and so reads no
    // annotation. The annotation cascade is therefore unreachable from exportCsv(), and these tests are the only
    // thing holding it up until the workbook form arrives.
    // ------------------------------------------------------------------

    @Test
    public void makeExportWorkbookMetaForCsv_withoutClassOrOption_holdsNoWorkbookAndTheCsvFormat() throws Exception {
        final PxlExportWorkbookMeta workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(null, null);

        // No POI workbook is created on this path, and neither is a formula evaluator. That is a complete state,
        // not a deficient one, which is what lets the CSV core share the metadata layer with the Excel core.
        assertThat(workbookMeta.getWorkbook()).isNull();
        assertThat(workbookMeta.getFormulaEvaluator()).isNull();
        assertThat(workbookMeta.getExportExcelEngine()).isNull();

        // The format is stated directly rather than derived from an engine, and it is what brings the CSV limits.
        assertThat(workbookMeta.getExportFileFormat()).isEqualTo(PxlFileFormat.CSV);
        assertThat(workbookMeta.getExportFileFormat().getMaxExportRows()).isEqualTo(PxlConstants.EXPORT_MAX_NUMBER_OF_CSV_ROWS);
        assertThat(workbookMeta.getExportFileFormat().getMaxExportColumns()).isEqualTo(PxlConstants.EXPORT_MAX_NUMBER_OF_CSV_COLUMNS);

        assertThat(workbookMeta.getExportCsvCharset()).isEqualTo(PxlConstants.DEFAULT_EXPORT_CSV_CHARSET);
        assertThat(workbookMeta.getExportCsvDelimiter()).isEqualTo(PxlConstants.DEFAULT_EXPORT_CSV_DELIMITER);
        assertThat(workbookMeta.isExportCsvBom()).isEqualTo(PxlConstants.DEFAULT_EXPORT_CSV_BOM);
    }

    @Test
    public void makeExportWorkbookMetaForCsv_annotatedClass_readsTheDeclaredCsvValues() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(CsvExportWorkbook.class, null);

        assertThat(workbookMeta.getExportCsvCharset()).isEqualTo("EUC-KR");
        assertThat(workbookMeta.getExportCsvDelimiter()).isEqualTo(';');
        assertThat(workbookMeta.isExportCsvBom()).isTrue();
    }

    @Test
    public void makeExportWorkbookMetaForCsv_unspecifiedAnnotation_fallsBackToTheDefaults() throws Exception {
        // Both annotation elements hold the sentinel, which must not be handed on as a usable value: an empty
        // charset name would reach Charset.forName, and NUL would reach the delimiter.
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(DefaultCsvWorkbook.class, null);

        assertThat(workbookMeta.getExportCsvCharset()).isEqualTo(PxlConstants.DEFAULT_EXPORT_CSV_CHARSET);
        assertThat(workbookMeta.getExportCsvDelimiter()).isEqualTo(PxlConstants.DEFAULT_EXPORT_CSV_DELIMITER);
    }

    @Test
    public void makeExportWorkbookMetaForCsv_option_beatsTheAnnotation() throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportCsvCharset("UTF-16BE")
                .exportCsvDelimiter('|')
                .exportCsvBom(false)
                .build();

        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(CsvExportWorkbook.class, option);

        assertThat(workbookMeta.getExportCsvCharset()).isEqualTo("UTF-16BE");
        assertThat(workbookMeta.getExportCsvDelimiter()).isEqualTo('|');
        assertThat(workbookMeta.isExportCsvBom()).isFalse();
    }

    @Test
    public void makeExportSheetMetas_csvValues_cascadeFromTheSheetAnnotationToTheWorkbook() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(CsvExportWorkbook.class, null);

        final List<PxlExportSheetMeta> sheetMetas =
                PxlExportSheetMeta.makeExportSheetMetas(CsvExportWorkbook.class, workbookMeta, null, false);

        final PxlExportSheetMeta cities = sheetMetaNamed(sheetMetas, "Cities");
        final PxlExportSheetMeta departments = sheetMetaNamed(sheetMetas, "Departments");

        // A CSV workbook is written as one file per sheet, so a sheet may depart from the workbook's charset.
        assertThat(cities.getExportCsvCharset()).isEqualTo("UTF-16LE");
        assertThat(cities.getExportCsvDelimiter()).isEqualTo('\t');

        // A sheet that states none of them inherits the workbook's values rather than the built-in defaults.
        assertThat(departments.getExportCsvCharset()).isEqualTo("EUC-KR");
        assertThat(departments.getExportCsvDelimiter()).isEqualTo(';');
    }

    @Test
    public void makeExportSheetMetas_csvBom_lettingASheetTurnOffWhatTheWorkbookAskedFor() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(CsvExportWorkbook.class, null);
        assertThat(workbookMeta.isExportCsvBom()).isTrue();

        final List<PxlExportSheetMeta> sheetMetas =
                PxlExportSheetMeta.makeExportSheetMetas(CsvExportWorkbook.class, workbookMeta, null, false);

        // This is the whole reason the annotation element is three-valued: a boolean one could not tell "turn the
        // workbook's mark off" apart from "say nothing", so the sheet could never get back to false.
        assertThat(sheetMetaNamed(sheetMetas, "Cities").isExportCsvBom()).isFalse();
        assertThat(sheetMetaNamed(sheetMetas, "Departments").isExportCsvBom()).isTrue();
    }

    @Test
    public void makeExportSheetMetas_csvBomSheetOption_beatsTheSheetAnnotation() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(CsvExportWorkbook.class, null);

        final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                .fieldName("cities")
                .exportCsvBom(true)     // the annotation says FALSE
                .build();

        final List<PxlExportSheetMeta> sheetMetas = PxlExportSheetMeta.makeExportSheetMetas(
                CsvExportWorkbook.class, workbookMeta, Arrays.asList(sheetOption), false);

        assertThat(sheetMetaNamed(sheetMetas, "Cities").isExportCsvBom()).isTrue();
    }

    @Test
    public void triState_unsetTellsNothingApartFromFalse() {
        assertThat(PxlOptionalBoolean.UNSPECIFIED.isSpecified()).isFalse();
        assertThat(PxlOptionalBoolean.TRUE.isSpecified()).isTrue();
        assertThat(PxlOptionalBoolean.FALSE.isSpecified()).isTrue();

        // UNSPECIFIED answers false too, which is exactly why isSpecified() has to be consulted first.
        assertThat(PxlOptionalBoolean.UNSPECIFIED.toBoolean()).isFalse();
        assertThat(PxlOptionalBoolean.TRUE.toBoolean()).isTrue();
        assertThat(PxlOptionalBoolean.FALSE.toBoolean()).isFalse();
    }

    @Test
    public void makeExportSheetMetas_csvSheetOption_beatsTheSheetAnnotation() throws Exception {
        final PxlExportWorkbookMeta workbookMeta =
                PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(CsvExportWorkbook.class, null);

        final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                .fieldName("cities")
                .exportCsvCharset("UTF-8")
                .exportCsvDelimiter('|')
                .build();

        final List<PxlExportSheetMeta> sheetMetas = PxlExportSheetMeta.makeExportSheetMetas(
                CsvExportWorkbook.class, workbookMeta, Arrays.asList(sheetOption), false);

        final PxlExportSheetMeta cities = sheetMetaNamed(sheetMetas, "Cities");
        assertThat(cities.getExportCsvCharset()).isEqualTo("UTF-8");
        assertThat(cities.getExportCsvDelimiter()).isEqualTo('|');
    }

    private static PxlExportSheetMeta sheetMetaNamed(final List<PxlExportSheetMeta> sheetMetas, final String name) {
        return sheetMetas.stream()
                .filter(sheetMeta -> name.equals(sheetMeta.getActualExportSheetName()))
                .findFirst()
                .orElseThrow(() -> new IllegalStateException("no sheet meta named " + name));
    }

    // ------------------------------------------------------------------
    // internal/core - CSV exporter
    // ------------------------------------------------------------------

    @Test
    public void coreCsvExporter_invalidDelimiter_isNormalizedRatherThanLeakingFromCommonsCsv() throws Exception {
        // The builder rejects an unusable delimiter before the destination is opened, so this backstop is only
        // reachable by calling the core directly - which the workbook form will do once it exists. Without it the
        // failure would surface as an unclassified system error naming nothing.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportCsvDelimiter('\n')
                .build();
        final PxlExportWorkbookMeta workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(null, option);

        assertThrows(PxlArgumentException.class, () -> PxlCoreCsvExporter.writeCsv(
                "Employees", Arrays.asList(new Employee()), Employee.class, workbookMeta, null, new StringWriter()));
    }

    @Test
    public void coreCsvExporter_missingArguments_areRejected() throws Exception {
        final PxlExportWorkbookMeta workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(null, null);
        final List<Employee> rows = Arrays.asList(new Employee());

        assertThrows(PxlArgumentException.class, () -> PxlCoreCsvExporter.writeCsv(
                " ", rows, Employee.class, workbookMeta, null, new StringWriter()));
        assertThrows(PxlNullPointerException.class, () -> PxlCoreCsvExporter.writeCsv(
                "S", null, Employee.class, workbookMeta, null, new StringWriter()));
        assertThrows(PxlNullPointerException.class, () -> PxlCoreCsvExporter.writeCsv(
                "S", rows, Employee.class, workbookMeta, null, null));
        assertThrows(PxlNullPointerException.class, () -> PxlCoreCsvExporter.writeSampleCsv(
                "S", null, workbookMeta, new StringWriter()));
    }

}
