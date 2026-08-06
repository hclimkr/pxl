package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlI18nException;
import io.github.hclimkr.pxl.exception.PxlValidationException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18n;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nContent;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.*;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.*;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.*;

/**
 * i18n (import/exportI18nBaseName, Language) verification.
 * <p>
 * Verifies that, via messages.properties (staff.column.role=Role, staff.column.fullName=Full Name, staff.sheet=Staff),
 * sheet/column names are translated on export and re-matched by the translated header/sheet names on import.
 */
public class PxlI18nTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    // Captures the current test method name to match it with the export file name.
    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

    private static I18nWorkbook sampleWorkbook() {
        final I18nRow row = new I18nRow();
        row.setRole("admin");
        row.setFullName("Alice");

        final I18nWorkbook workbook = new I18nWorkbook();
        workbook.setWorkbookName("W");
        workbook.setPeople(Arrays.asList(row));
        return workbook;
    }

    // ------------------------------------------------------------------
    // export: sheet/column names are translated when written
    // ------------------------------------------------------------------

    @Test
    public void exportI18n_bundle_translatesSheetAndHeaders() throws Exception {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .workbook(sampleWorkbook())
                .override(noValidationOption())
                .toStream(outputStream);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            // sheet key staff.sheet -> "Staff"
            final Sheet sheet = workbook.getSheet("Staff");
            assertThat(sheet).as("sheet name should be translated to 'Staff'").isNotNull();

            // column keys staff.column.role/fullName -> Role/Full Name
            final Row header = sheet.getRow(0);
            final Set<String> headers = new HashSet<>();
            for (final Cell cell : header) {
                headers.add(cell.getStringCellValue());
            }
            assertThat(headers).contains("Role", "Full Name");
        }
    }

    // ------------------------------------------------------------------
    // import: re-matched by translated header/sheet names (round-trip)
    // ------------------------------------------------------------------

    @Test
    public void importI18n_translatedNames_roundTrips() throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(sampleWorkbook())
                .override(noValidationOption())
                .toFile(excelFile);

        final I18nWorkbook imported = pxl.importExcel()
                .workbookName("W")
                .workbook(I18nWorkbook.class)
                .fromFile(excelFile);

        assertThat(imported.getPeople()).as("should match via translated sheet/header names").hasSize(1);
        final I18nRow row = imported.getPeople().get(0);
        assertThat(row.getRole()).isEqualTo("admin");
        assertThat(row.getFullName()).isEqualTo("Alice");
    }

    // ------------------------------------------------------------------
    // The names an option overrides are bundle keys too, just like the annotation names they replace
    // ------------------------------------------------------------------

    private static PxlExportWorkbookOption overriddenNameExportOption() {
        final PxlExportColumnOption columnOption = PxlExportColumnOption.builder()
                .fieldName("role")
                .exportColumnNames(Arrays.asList("staff.override.column"))
                .build();
        final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                .fieldName("people")
                .exportSheetNames(Arrays.asList("staff.override.sheet"))
                .exportColumnOptions(Arrays.asList(columnOption))
                .build();

        return PxlExportWorkbookOption.builder()
                .exportSheetOptions(Arrays.asList(sheetOption))
                .exportDataValidation(false)
                .build();
    }

    private static PxlImportWorkbookOption overriddenNameImportOption() {
        final PxlImportColumnOption columnOption = PxlImportColumnOption.builder()
                .fieldName("role")
                .importColumnNames(Arrays.asList("staff.override.column"))
                .build();
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .fieldName("people")
                .importSheetNames(Arrays.asList("staff.override.sheet"))
                .importColumnOptions(Arrays.asList(columnOption))
                .build();

        return PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();
    }

    @Test
    public void exportI18n_optionOverriddenNames_translatedThroughBundle() throws Exception {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .workbook(sampleWorkbook())
                .override(overriddenNameExportOption())
                .toStream(outputStream);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            // staff.override.sheet -> "Overridden Staff": the overriding name goes through the bundle as well
            final Sheet sheet = workbook.getSheet("Overridden Staff");
            assertThat(sheet).as("the sheet name supplied by the option should be translated too").isNotNull();

            final Row header = sheet.getRow(0);
            final Set<String> headers = new HashSet<>();
            for (final Cell cell : header) {
                headers.add(cell.getStringCellValue());
            }
            // the overridden column is translated from staff.override.column, the untouched one from its annotation name
            assertThat(headers).contains("Overridden Role", "Full Name");
        }
    }

    @Test
    public void importI18n_optionOverriddenNames_matchThroughBundle() throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(sampleWorkbook())
                .override(overriddenNameExportOption())
                .toFile(excelFile);

        // The import side resolves the overriding names through the bundle as well, so the translated
        // sheet/header written above are found again.
        final I18nWorkbook imported = pxl.importExcel()
                .workbookName("W")
                .override(overriddenNameImportOption())
                .workbook(I18nWorkbook.class)
                .fromFile(excelFile);

        assertThat(imported.getPeople()).as("should match via the translated overriding names").hasSize(1);
        assertThat(imported.getPeople().get(0).getRole()).isEqualTo("admin");
        assertThat(imported.getPeople().get(0).getFullName()).isEqualTo("Alice");
    }

    // ------------------------------------------------------------------
    // Country: language+country (ko_KR) -> resolves to the Korean bundle (messages_ko)
    // ------------------------------------------------------------------

    private static I18nCountryWorkbook sampleCountryWorkbook() {
        final I18nRow row = new I18nRow();
        row.setRole("admin");
        row.setFullName("Bob");

        final I18nCountryWorkbook workbook = new I18nCountryWorkbook();
        workbook.setWorkbookName("W");
        workbook.setPeople(Arrays.asList(row));
        return workbook;
    }

    @Test
    public void exportI18nCountry_bundle_usesCountryVariant() throws Exception {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .workbook(sampleCountryWorkbook())
                .override(noValidationOption())
                .toStream(outputStream);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            // language=ko, country=KR -> messages_ko.properties (Korean translation, different from the English base)
            final Sheet sheet = workbook.getSheet("직원");
            assertThat(sheet).as("sheet name should be the Korean translation").isNotNull();

            final Row header = sheet.getRow(0);
            final Set<String> headers = new HashSet<>();
            for (final Cell cell : header) {
                headers.add(cell.getStringCellValue());
            }
            assertThat(headers).contains("역할", "성명");
        }
    }

    @Test
    public void importI18nCountry_translatedNames_roundTrips() throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(sampleCountryWorkbook())
                .override(noValidationOption())
                .toFile(excelFile);

        final I18nCountryWorkbook imported =
                pxl.importExcel()
                        .workbookName("W")
                        .workbook(I18nCountryWorkbook.class)
                        .fromFile(excelFile);

        assertThat(imported.getPeople()).hasSize(1);
        assertThat(imported.getPeople().get(0).getRole()).isEqualTo("admin");
        assertThat(imported.getPeople().get(0).getFullName()).isEqualTo("Bob");
    }

    // ------------------------------------------------------------------
    // A bundle injected into the option wins over the annotated base name
    // ------------------------------------------------------------------

    // The English base bundle, read against a workbook whose annotation asks for Korean. Every value here is ASCII,
    // so the assertions below do not depend on how properties files happen to be decoded.
    private static ResourceBundle baseBundle() {
        return ResourceBundle.getBundle("messages", Locale.ROOT);
    }

    @Test
    public void exportI18n_optionResourceBundle_overridesAnnotationBaseName() throws Exception {
        // I18nCountryWorkbook declares ko_KR, so left alone it writes "직원"/"역할". Handing the option the base
        // bundle must displace that entirely - the annotated triple is not consulted at all.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportResourceBundle(baseBundle())
                .exportDataValidation(false)
                .build();

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .workbook(sampleCountryWorkbook())
                .override(option)
                .toStream(outputStream);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            assertThat(workbook.getSheet("Staff")).as("the injected bundle should decide the sheet name").isNotNull();
            assertThat(workbook.getSheet("직원")).as("the annotated ko_KR bundle should not be loaded").isNull();

            final Set<String> headers = new HashSet<>();
            for (final Cell cell : workbook.getSheet("Staff").getRow(0)) {
                headers.add(cell.getStringCellValue());
            }
            assertThat(headers).contains("Role", "Full Name");
        }
    }

    @Test
    public void importI18n_optionResourceBundle_overridesAnnotationBaseName() throws Exception {
        // Written with the English names above, the same workbook binds them back only because the import side
        // resolves its names through the injected bundle too - its own annotation would look for the Korean ones.
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(sampleCountryWorkbook())
                .override(PxlExportWorkbookOption.builder()
                        .exportResourceBundle(baseBundle())
                        .exportDataValidation(false)
                        .build())
                .toFile(excelFile);

        final I18nCountryWorkbook imported = pxl.importExcel()
                .workbookName("W")
                .override(PxlImportWorkbookOption.builder()
                        .importResourceBundle(baseBundle())
                        .build())
                .workbook(I18nCountryWorkbook.class)
                .fromFile(excelFile);

        assertThat(imported.getPeople()).as("should match via the injected bundle's names").hasSize(1);
        assertThat(imported.getPeople().get(0).getRole()).isEqualTo("admin");
        assertThat(imported.getPeople().get(0).getFullName()).isEqualTo("Bob");
    }

    @Test
    public void importI18n_optionResourceBundle_unknownAnnotationBaseName_doesNotThrow() throws Exception {
        // I18nMissingBundleWorkbook names a bundle that does not exist, which on its own is a PxlI18nException.
        // An injected bundle takes its place before it is ever loaded, so the import succeeds.
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(sampleCountryWorkbook())
                .override(PxlExportWorkbookOption.builder()
                        .exportResourceBundle(baseBundle())
                        .exportDataValidation(false)
                        .build())
                .toFile(excelFile);

        assertThatThrownBy(() -> pxl.importExcel()
                .workbook(I18nMissingBundleWorkbook.class)
                .fromFile(excelFile))
                .as("without a bundle the missing base name should fail")
                .isInstanceOf(PxlI18nException.class);

        final I18nMissingBundleWorkbook imported = pxl.importExcel()
                .override(PxlImportWorkbookOption.builder()
                        .importResourceBundle(baseBundle())
                        .build())
                .workbook(I18nMissingBundleWorkbook.class)
                .fromFile(excelFile);

        assertThat(imported.getPeople()).as("the injected bundle should replace the missing one").hasSize(1);
        assertThat(imported.getPeople().get(0).getFullName()).isEqualTo("Bob");
    }

    // ------------------------------------------------------------------
    // export sample: exportSample and exportOptionItems are translated as well
    // ------------------------------------------------------------------

    // Reads a sample sheet as a header-name -> sample-cell-text map. Rendered through DataFormatter so a numeric
    // sample cell reads back as text too.
    private static Map<String, String> samplesByHeader(final Sheet sheet) {
        final DataFormatter dataFormatter = new DataFormatter();
        final Row header = sheet.getRow(0);
        final Row sampleRow = sheet.getRow(1);

        final Map<String, String> samples = new HashMap<>();
        for (final Cell headerCell : header) {
            final Cell sampleCell = sampleRow.getCell(headerCell.getColumnIndex());
            samples.put(headerCell.getStringCellValue(), dataFormatter.formatCellValue(sampleCell));
        }
        return samples;
    }

    @Test
    public void exportSampleI18n_scalarAndCollection_translatesEveryElement() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(I18nSampleWorkbook.class)
                .toWorkbook();
        try {
            final Sheet sheet = workbook.getSheet("Staff");
            assertThat(sheet).as("sheet name should be translated to 'Staff'").isNotNull();

            final Map<String, String> samples = samplesByHeader(sheet);

            // A String sample is a bundle key, exactly like a column name.
            assertThat(samples.get("Role")).isEqualTo("Administrator");
            // A Collection sample carries one key per element, so each element is translated on its own
            // and they are joined back with the collection separator.
            assertThat(samples.get("Roles")).isEqualTo("Administrator;User");
            // Collection<Enum>: each element is translated first, then parsed back into its constant, so the
            // cell holds the canonical names.
            assertThat(samples.get("Grades")).isEqualTo("A;B");
            assertThat(samples.get("Grade")).isEqualTo("A");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSampleI18n_customCollectionSeparator_splitsOnTheColumnSeparator() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(I18nSampleWorkbook.class)
                .toWorkbook();
        try {
            final Map<String, String> samples = samplesByHeader(workbook.getSheet("Staff"));

            // The separator comes from exportCollectionSeparator ("::"), not from the bundle value, so changing it
            // keeps working without touching the bundle.
            assertThat(samples.get("Tags")).isEqualTo("Administrator::User");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportSampleI18n_nonStringColumn_keepsSampleAndUnknownHeaderVerbatim() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(I18nSampleWorkbook.class)
                .toWorkbook();
        try {
            final Map<String, String> samples = samplesByHeader(workbook.getSheet("Staff"));

            // "count" is not in the bundle, so the header passes through untranslated.
            assertThat(samples).containsKey("count");
            // The sample of a numeric column is never translated: "1234" is a bundle key (1234=9999), and applying it
            // would put a value the column never declared into the cell.
            assertThat(samples.get("count")).isEqualTo("1234");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportOptionItemsI18n_stringColumnTranslated_enumColumnVerbatim() throws Exception {
        final Workbook workbook = pxl.exportSampleExcel()
                .workbook(I18nSampleWorkbook.class)
                .toWorkbook();
        try {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Staff");

            final List<List<String>> optionLists = new ArrayList<>();
            for (final XSSFDataValidation dataValidation : sheet.getDataValidations()) {
                optionLists.add(Arrays.asList(dataValidation.getValidationConstraint().getExplicitListValues()));
            }

            // String column: the dropdown offers the very text the sample cell holds, so the sample does not
            // violate its own data validation.
            assertThat(optionLists).contains(Arrays.asList("Administrator", "User"));
            // Enum column: the cell always holds the canonical constant, so the items are used as declared.
            assertThat(optionLists).contains(Arrays.asList("staff.grade.a", "staff.grade.b"));
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Diagnostic-message i18n (exception text) - a global locale independent of workbook content i18n.
    // The default is the JVM default locale, with a process-wide override via Pxl.setMessageLocale.
    // The base bundle is English and Korean (_ko) exists; an unmatched locale falls back to English.
    // Since this is global state, always revert with resetMessageLocale in finally.
    // ------------------------------------------------------------------

    @Test
    public void messageLocale_english_tagRendersInEnglish() {
        try {
            Pxl.setMessageLocale(Locale.ENGLISH);

            // rowIndex is passed 0-based and rendered 1-based (3) in the message. When columnName is present, columnIndex is ignored.
            final PxlValidationException exception = new PxlValidationException("Users", 2, "age", null, "boom");

            assertThat(exception.getMessage()).isEqualTo("sheet 'Users', row 3, column 'age': boom");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_korean_tagRendersInKorean() {
        try {
            Pxl.setMessageLocale(Locale.KOREAN);

            final PxlValidationException exception = new PxlValidationException("Users", 2, "age", null, "boom");

            assertThat(exception.getMessage()).isEqualTo("'Users' 시트, 3행, 'age' 열: boom");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_unmatchedLocale_fallsBackToEnglishBase() {
        try {
            // A locale without a bundle (fr) must fall back to the base (English).
            Pxl.setMessageLocale(Locale.FRENCH);

            final PxlValidationException exception = new PxlValidationException("Users", 2, "age", null, "boom");

            assertThat(exception.getMessage()).isEqualTo("sheet 'Users', row 3, column 'age': boom");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_assertPrecondition_localizesNamedArgument() {
        try {
            Pxl.setMessageLocale(Locale.ENGLISH);
            assertThatThrownBy(() -> PxlAssertSupport.notNull(null, "rowClass"))
                    .hasMessage("argument 'rowClass' is null.");

            Pxl.setMessageLocale(Locale.KOREAN);
            assertThatThrownBy(() -> PxlAssertSupport.notNull(null, "rowClass"))
                    .hasMessage("'rowClass' 인자가 null입니다.");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_parameterizedMessage_substitutesParamsPerLocale() {
        try {
            // codec.parse.invalid: {0}=value, {1}=type. Verifies MessageFormat substitution and quote escaping ('') in both locales.
            Pxl.setMessageLocale(Locale.ENGLISH);
            assertThat(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, "abc", "Integer"))
                    .isEqualTo("'abc' is not a valid Integer value.");

            Pxl.setMessageLocale(Locale.KOREAN);
            assertThat(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, "abc", "Integer"))
                    .isEqualTo("'abc'은(는) 올바른 형식의 Integer 값이 아닙니다.");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_convertUnsupported_rendersQuotesAndJosa() {
        try {
            // codec.convert.unsupported: verifies the Korean particle literal '(으)로' and quotes are rendered as-is.
            Pxl.setMessageLocale(Locale.ENGLISH);
            assertThat(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, "String", "LocalDate"))
                    .isEqualTo("a value of type 'String' cannot be converted to LocalDate.");

            Pxl.setMessageLocale(Locale.KOREAN);
            assertThat(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, "String", "LocalDate"))
                    .isEqualTo("'String' 타입의 값을 LocalDate(으)로 변환할 수 없습니다.");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_numericParam_hasNoThousandsGrouping() {
        try {
            // Numeric parameters are passed via String.valueOf so MessageFormat's digit grouping (1,048,576) must not be applied.
            Pxl.setMessageLocale(Locale.ENGLISH);
            final String message = PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(1048576));

            assertThat(message).isEqualTo("the number of sheets exceeds the limit of 1048576.");
            assertThat(message).doesNotContain("1,048,576");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_missingKey_returnsKeyAndNeverThrows() {
        try {
            Pxl.setMessageLocale(Locale.ENGLISH);

            // fail-safe: a nonexistent key returns the key string as-is without throwing (so exception assembly does not mask the original error).
            assertThatCode(() -> PxlI18nDiagnostic.get("pxl.nonexistent.key", "ignored"))
                    .doesNotThrowAnyException();
            assertThat(PxlI18nDiagnostic.get("pxl.nonexistent.key", "ignored"))
                    .isEqualTo("pxl.nonexistent.key");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void messageLocale_reset_revertsToJvmDefaultLocale() {
        final Locale savedDefault = Locale.getDefault();
        try {
            // Clearing the override (reset) follows the JVM default locale. Verified by changing the JVM default.
            Pxl.resetMessageLocale();

            Locale.setDefault(Locale.KOREAN);
            assertThat(new PxlValidationException("S", 0, "c", null, "boom").getMessage())
                    .isEqualTo("'S' 시트, 1행, 'c' 열: boom");

            Locale.setDefault(Locale.ENGLISH);
            assertThat(new PxlValidationException("S", 0, "c", null, "boom").getMessage())
                    .isEqualTo("sheet 'S', row 1, column 'c': boom");
        } finally {
            Locale.setDefault(savedDefault);
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void exportExcel_missingConfig_throwsLocalizedMessageThroughApi() {
        // Throw a real exception through the public API to verify the whole chain (throw -> terminal PxlException wrapping -> localized message) works.
        // The terminal wraps in PxlException(cause) so the message gets a class prefix; verified with hasMessageContaining.
        try {
            Pxl.setMessageLocale(Locale.ENGLISH);
            assertThatThrownBy(() -> pxl.exportExcel()
                    .toStream(new ByteArrayOutputStream()))
                    .isInstanceOf(PxlArgumentException.class)
                    .hasMessageContaining("either workbook(Object) or sheet(...) must be specified.");

            Pxl.setMessageLocale(Locale.KOREAN);
            assertThatThrownBy(() -> pxl.exportExcel()
                    .toStream(new ByteArrayOutputStream()))
                    .isInstanceOf(PxlArgumentException.class)
                    .hasMessageContaining("workbook(Object) 또는 sheet(...) 중 하나를 지정해야 합니다.");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    // ------------------------------------------------------------------
    // Boundary behavior of the shared loader (PxlI18n) and the content facade (PxlI18nContent).
    // PxlI18nException if a bundle is specified but missing; null meaning "i18n disabled" if baseName/locale is unspecified;
    // a translation gap (null bundle or missing key) passes the name through unchanged.
    // ------------------------------------------------------------------

    @Test
    public void getBundle_configuredBundleMissing_throwsPxlI18nException() {
        // The loader throws PxlI18nException when baseName is specified but the resource is missing (shared path for content and diagnostic channels).
        // This bootstrap failure message is not localized but hardcoded in English to prevent recursion.
        assertThatThrownBy(() -> PxlI18n.getBundle("pxl-no-such-bundle", Locale.ENGLISH))
                .isInstanceOf(PxlI18nException.class)
                .hasMessageContaining("cannot be found");
    }

    @Test
    public void getBundle_i18nDisabledInputs_returnNullWithoutThrowing() throws Exception {
        // baseName blank / locale null / language or country null -> treated as i18n disabled, returning null without throwing.
        // (language and country are disabled only when null, not when blank.)
        assertThat(PxlI18n.getBundle("", Locale.ENGLISH)).isNull();
        assertThat(PxlI18n.getBundle("messages", (Locale) null)).isNull();
        assertThat(PxlI18n.getBundle("", "en", "US")).isNull();
        assertThat(PxlI18n.getBundle("messages", null, "US")).isNull();
        assertThat(PxlI18n.getBundle("messages", "en", null)).isNull();
    }

    @Test
    public void getBundle_emptyLanguageOrCountry_treatedAsProvidedNotDisabled() throws Exception {
        // Changed contract: language/country are disabled only when null. An empty string is treated as 'provided',
        // so the locale composition proceeds (without throwing) and resolves to the base bundle.
        assertThat(PxlI18n.getBundle("messages", "en", "")).isNotNull();
        assertThat(PxlI18n.getBundle("messages", "", "US")).isNotNull();
    }

    @Test
    public void contentLoadBundle_disabledInputs_returnNull() throws Exception {
        // The content facade is also disabled -> null when baseName is blank or language/country is null.
        assertThat(PxlI18nContent.loadBundle("", "en", "US")).isNull();
        assertThat(PxlI18nContent.loadBundle("messages", null, "US")).isNull();
        assertThat(PxlI18nContent.loadBundle("messages", "en", null)).isNull();
    }

    @Test
    public void contentTranslate_disabledOrMissingKey_returnsNameUnchanged() throws Exception {
        // When the bundle is null (content i18n disabled), the name is returned unchanged.
        assertThat(PxlI18nContent.translate(null, "staff.column.role")).isEqualTo("staff.column.role");

        final ResourceBundle bundle = PxlI18nContent.loadBundle("messages", "ko", "KR");
        assertThat(bundle).as("the messages_ko bundle should be loaded").isNotNull();
        // An existing key is translated, and a missing key silently passes the original name through.
        assertThat(PxlI18nContent.translate(bundle, "staff.column.role")).isEqualTo("역할");
        assertThat(PxlI18nContent.translate(bundle, "nonexistent")).isEqualTo("nonexistent");
    }

    @Test
    public void contentBundle_unmatchedLocale_doesNotFallBackToJvmDefault() throws Exception {
        final Locale savedDefault = Locale.getDefault();
        try {
            // messages_ko exists and is made the JVM default, so a bundle asked for in French would land on the
            // Korean one if the loader kept the JVM locale fallback.
            Locale.setDefault(Locale.KOREAN);

            final ResourceBundle bundle = PxlI18nContent.loadBundle("messages", "fr", "");
            assertThat(bundle).as("an unmatched locale should still resolve to the base bundle").isNotNull();

            // Utf8ResourceControl#getFallbackLocale returns null, so fr resolves to the English base, not messages_ko.
            assertThat(PxlI18nContent.translate(bundle, "staff.column.role")).isEqualTo("Role");
        } finally {
            Locale.setDefault(savedDefault);
        }
    }
}
