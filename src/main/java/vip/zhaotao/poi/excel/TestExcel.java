package vip.zhaotao.poi.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import vip.zhaotao.poi.annotation.ExcelColumn;

import java.math.BigDecimal;
import java.util.Date;

/**
 * Test excel
 *
 * @author zhaotao
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class TestExcel {

    @ExcelColumn(name = "Character", number = 0)
    private Character characterValue;

    @ExcelColumn(name = "String", number = 1)
    private String stringValue;

    @ExcelColumn(name = "Byte", number = 2)
    private Byte byteValue;

    @ExcelColumn(name = "Short", number = 3)
    private Short shortValue;

    @ExcelColumn(name = "Integer", number = 4, format = "0%")
    private Integer integerValue;

    @ExcelColumn(name = "Long", number = 5, format = "[DbNum2][$-804]0")
    private Long longValue;

    @ExcelColumn(name = "Float", number = 6, format = "0.00")
    private Float floatValue;

    @ExcelColumn(name = "Double", number = 7, format = "¥#,##0.00;¥-#,##0.00")
    private Double doubleValue;

    @ExcelColumn(name = "Date", number = 8, format = "yyyy-MM-dd HH:mm:ss")
    private Date dateValue;

    @ExcelColumn(name = "Boolean", number = 9)
    private Boolean booleanValue;

    @ExcelColumn(name = "BigDecimal", number = 10)
    private BigDecimal bigDecimalValue;
}
