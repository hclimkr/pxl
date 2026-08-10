package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.type.PxlOptionalBoolean;
import lombok.*;

import java.util.List;

/**
 * A workbook declaring the CSV export charset/delimiter/BOM at both annotation levels.
 * <p>
 * CSV export has the sheet form only, which builds its metadata with no workbook class and therefore reads no
 * annotation at all, so these elements are exercised through the metadata factories rather than through
 * {@code exportCsv()}. They are declared now so that adding the workbook form later activates them without having
 * to change any default.
 * <ul>
 *   <li>{@code cities} states all three, so the sheet level must win over the workbook level. Its BOM says
 *       {@link PxlOptionalBoolean#FALSE}, which is the case a plain {@code boolean} element could not express: it has to
 *       turn off a mark the workbook asked for rather than read as "not specified".</li>
 *   <li>{@code departments} states none of them, so it must inherit the workbook level.</li>
 * </ul>
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportCsvCharset = "EUC-KR", exportCsvDelimiter = ';', exportCsvBom = PxlOptionalBoolean.TRUE)
public class CsvExportWorkbook {

    @PxlSheet(name = "Cities", exportCsvCharset = "UTF-16LE", exportCsvDelimiter = '\t', exportCsvBom = PxlOptionalBoolean.FALSE)
    private List<CharsetRow> cities;

    @PxlSheet(name = "Departments")
    private List<Department> departments;

}
