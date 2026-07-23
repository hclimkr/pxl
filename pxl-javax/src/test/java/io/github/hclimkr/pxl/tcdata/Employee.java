package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.math.BigDecimal;
import java.time.LocalDate;

/**
 * A simple row DTO used broadly across multi-sheet/workbook/CSV/grouping tests.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Employee {

    @PxlColumn(name = "Name", exportSample = "Alice")
    private String name;

    @PxlColumn(name = "Age", exportSample = "30")
    private int age;

    @PxlColumn(name = "Salary", exportSample = "50000.50")
    private BigDecimal salary;

    @PxlColumn(name = "Active", exportSample = "true")
    private Boolean active;

    @PxlColumn(name = "HireDate", pattern = "yyyy-MM-dd", exportSample = "2020-01-15")
    private LocalDate hireDate;

    @PxlColumn(name = "Grade", exportSample = "A")
    private Grade grade;

    // The grouping (exportGroupingFieldName) key field
    @PxlColumn(name = "Department", exportSample = "Engineering")
    private String department;

}
