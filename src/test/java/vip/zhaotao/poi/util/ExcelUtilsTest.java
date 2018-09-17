package vip.zhaotao.poi.util;

import com.google.common.collect.Lists;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.RandomUtils;
import org.junit.Test;
import vip.zhaotao.poi.excel.TestExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

@Slf4j
public class ExcelUtilsTest {

    @Test
    @SneakyThrows
    public void write() {
        String filePath = this.getFilePath();
        FileOutputStream outputStream = new FileOutputStream(filePath);
        ExcelUtils.write(outputStream, this.getTestData());
        String command = String.format("cmd /c start excel %s", filePath);
        Runtime.getRuntime().exec(command);
    }

    @Test
    @SneakyThrows
    public void read() {
        FileInputStream inputStream = new FileInputStream(this.getFilePath());
        List<TestExcel> userExcelList = ExcelUtils.read(inputStream, TestExcel.class);
        log.info(userExcelList.toString());
    }

    public String getFilePath() {
        StringBuilder filePath = new StringBuilder();
        filePath.append(System.getProperty("user.home").replace("\\", "/"));
        if (System.getProperty("os.name").startsWith("Windows")) {
            filePath.append("/Desktop");
        }
        filePath.append("/test").append(ExcelUtils.Type.OFFICE_OPEN_XML_SHEET.getExtensionName());
        return filePath.toString();
    }

    private List<TestExcel> getTestData() {
        List<TestExcel> list = Lists.newArrayList();
        TestExcel testExcel = new TestExcel();
        testExcel.setCharacterValue('A');
        testExcel.setStringValue(RandomStringUtils.randomAlphanumeric(1, 30));
        testExcel.setByteValue(Byte.MIN_VALUE);
        testExcel.setShortValue(Short.MIN_VALUE);
        testExcel.setIntegerValue(RandomUtils.nextInt(1, 10000));
        testExcel.setLongValue(RandomUtils.nextLong(1, 10000));
        testExcel.setFloatValue(RandomUtils.nextFloat(1f, 10000f));
        testExcel.setDoubleValue(RandomUtils.nextDouble(1d, 10000d));
        testExcel.setDateValue(new Date());
        testExcel.setBooleanValue(RandomUtils.nextBoolean());
        testExcel.setBigDecimalValue(BigDecimal.ONE);
        list.add(testExcel);
        return list;
    }
}