package vip.zhaotao.poi.util;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Lombok;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.CharUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import vip.zhaotao.poi.annotation.ExcelColumn;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Excel util
 *
 * @author zhaotao
 */
@Slf4j
public class ExcelUtils {

    /**
     * Default start row number
     */
    private static final Integer DEFAULT_START_ROW_NUM = BigInteger.ONE.intValue();

    /**
     * Default column width
     */
    private static final Integer DEFAULT_COLUMN_WIDTH = BigInteger.TEN.intValue();

    public static <T> void write(OutputStream outputStream, List<T> dataList) {
        write(Type.OFFICE_OPEN_XML_SHEET, outputStream, dataList);
    }

    public static <T> void write(Type type, OutputStream outputStream, List<T> dataList) {
        write(type, outputStream, dataList, DEFAULT_COLUMN_WIDTH);
    }

    /**
     * Write file
     *
     * @param type
     * @param outputStream
     * @param dataList
     * @param defaultColumnWidth
     * @param <T>
     */
    public static <T> void write(Type type, OutputStream outputStream, List<T> dataList, Integer defaultColumnWidth) {
        if (type == null || outputStream == null || CollectionUtils.isEmpty(dataList)) {
            return;
        }
        Workbook workbook = null;
        try {
            workbook = getWorkbook(type);
            CreationHelper creationHelper = workbook.getCreationHelper();
            Sheet sheet = workbook.createSheet();
            sheet.setDefaultColumnWidth(defaultColumnWidth);
            Class<?> clazz = dataList.get(0).getClass();
            createHeader(sheet, clazz);

            TreeMap<Integer, ExcelColumnAnnotationInfo> columnNumberMap = getColumnNumberMap(clazz);
            Iterator<Map.Entry<Integer, ExcelColumnAnnotationInfo>> iterator = columnNumberMap.entrySet().iterator();
            for (T t : dataList) {
                int lastRowNum = sheet.getLastRowNum();
                Row row = sheet.createRow(++lastRowNum);
                while (iterator.hasNext()) {
                    Map.Entry<Integer, ExcelColumnAnnotationInfo> next = iterator.next();
                    Integer key = next.getKey();
                    ExcelColumnAnnotationInfo value = next.getValue();
                    String format = value.getFormat();
                    Object fieldValue = FieldUtils.readField(value.getField(), t, true);
                    if (fieldValue == null) {
                        continue;
                    }
                    Cell cell = row.createCell(key);
                    if (fieldValue instanceof Number) {
                        cell.setCellValue(((Number) fieldValue).doubleValue());
                    } else if (fieldValue instanceof Date) {
                        cell.setCellValue((Date) fieldValue);
                        value.setFormat(StringUtils.isBlank(format) ? String.format("%s %s", DateFormatUtils.ISO_8601_EXTENDED_DATE_FORMAT.getPattern(), DateFormatUtils.ISO_8601_EXTENDED_TIME_FORMAT.getPattern()) : format);
                        sheet.setColumnWidth(key, format.length() * 256);
                    } else if (fieldValue instanceof Boolean) {
                        cell.setCellValue((Boolean) fieldValue);
                    } else {
                        cell.setCellValue(fieldValue.toString());
                        int fieldValueLength = fieldValue.toString().getBytes(StandardCharsets.UTF_8.name()).length;
                        if (fieldValueLength > sheet.getDefaultColumnWidth()) {
                            sheet.setColumnWidth(key, fieldValueLength * 256);
                        }
                    }
                    if (StringUtils.isNotBlank(format)) {
                        CellStyle cellStyle = workbook.createCellStyle();
                        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(format));
                        cell.setCellStyle(cellStyle);
                    }
                }
            }
            workbook.write(outputStream);
        } catch (Throwable t) {
            throw Lombok.sneakyThrow(t);
        } finally {
            IOUtils.closeQuietly(workbook);
            IOUtils.closeQuietly(outputStream);
        }
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clazz) {
        return read(inputStream, null, DEFAULT_START_ROW_NUM, clazz);
    }

    public static <T> List<T> read(InputStream inputStream, String password, Class<T> clazz) {
        return read(inputStream, password, DEFAULT_START_ROW_NUM, clazz);
    }

    public static <T> List<T> read(InputStream inputStream, Integer startRowNum, Class<T> clazz) {
        return read(inputStream, null, startRowNum, clazz);
    }

    /**
     * Read file content
     *
     * @param inputStream
     * @param password
     * @param startRowNum apply to all sheet (0-based)
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> read(InputStream inputStream, String password, Integer startRowNum, Class<T> clazz) {
        List<T> dataList = Lists.newLinkedList();
        if (inputStream == null || clazz == null) {
            return dataList;
        }
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream, password);
            // Get all fields of the class, the key is the name of the ExcelColumn annotation
            HashMap<String, Field> columnNameMap = getClassColumnNameMap(clazz);
            // Sheet processing
            int sheetNumber = workbook.getNumberOfSheets();
            for (int i = 0; i < sheetNumber; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                // Key is the column number
                LinkedHashMap<Integer, Field> sheetColumnNumberFieldMap = getSheetColumnNumberFieldMap(columnNameMap, sheet);
                // Row processing
                int lastRowNum = sheet.getLastRowNum();
                for (int rowNum = startRowNum; rowNum <= lastRowNum; rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    T t = clazz.newInstance();
                    // Column processing
                    for (int colNum = row.getFirstCellNum(); colNum < row.getLastCellNum(); colNum++) {
                        if (!sheetColumnNumberFieldMap.containsKey(colNum)) {
                            throw new RuntimeException(String.format("%s column did not find header.", CellReference.convertNumToColString(colNum)));
                        }
                        Cell cell = row.getCell(colNum);
                        Object cellValue = getCellValue(cell);
                        if (cellValue == null) {
                            log.info("{} cell value is empty.", cell.getAddress().formatAsString());
                            continue;
                        }
                        Field field = sheetColumnNumberFieldMap.get(colNum);
                        Class<?> fieldType = field.getType();
                        cellValue = cellValueProcessing(cellValue, fieldType);
                        if (!cellValue.getClass().equals(fieldType)) {
                            throw new RuntimeException(String.format("%s cell data type value is invalid, expected type is %s, actual type is %s.",
                                    cell.getAddress().formatAsString(), fieldType.getSimpleName(), cellValue.getClass().getSimpleName()));
                        }
                        FieldUtils.writeField(field, t, cellValue, true);
                    }
                    dataList.add(t);
                }
            }
        } catch (Throwable t) {
            throw Lombok.sneakyThrow(t);
        } finally {
            IOUtils.closeQuietly(workbook);
            IOUtils.closeQuietly(inputStream);
        }
        return dataList;
    }

    private static Workbook getWorkbook(Type type) {
        Workbook workbook = null;
        switch (type) {
            case MICROSOFT_EXCEL:
                workbook = new HSSFWorkbook();
                break;
            case OFFICE_OPEN_XML_SHEET:
                workbook = new SXSSFWorkbook(1);
                break;
        }
        return workbook;
    }

    private static LinkedHashMap<Integer, Field> getSheetColumnNumberFieldMap(Map<String, Field> columnNameMap, Sheet sheet) {
        // Key is column number
        LinkedHashMap<Integer, Field> header = Maps.newLinkedHashMap();
        int firstRowNum = sheet.getFirstRowNum();
        Row firstRow = sheet.getRow(firstRowNum);
        for (int colNum = firstRow.getFirstCellNum(); colNum < firstRow.getLastCellNum(); colNum++) {
            Object cellValue = getCellValue(firstRow.getCell(colNum));
            if (!columnNameMap.containsKey(cellValue)) {
                throw new RuntimeException(String.format("Unknown header, %s.", cellValue));
            }
            header.put(colNum, columnNameMap.get(cellValue));
        }
        return header;
    }

    private static Object cellValueProcessing(Object cellValue, Class<?> fieldType) {
        Object result;
        String cellValueString = cellValue.toString();
        if (fieldType.equals(Byte.class)) {
            result = NumberUtils.toByte(cellValueString);
        } else if (fieldType.equals(Short.class)) {
            result = NumberUtils.toShort(cellValueString);
        } else if (fieldType.equals(Integer.class)) {
            result = NumberUtils.toInt(cellValueString);
        } else if (fieldType.equals(Long.class)) {
            result = NumberUtils.toLong(cellValueString);
        } else if (fieldType.equals(Float.class)) {
            result = NumberUtils.toFloat(cellValueString);
        } else if (fieldType.equals(Double.class)) {
            result = NumberUtils.toDouble(cellValueString);
        } else if (fieldType.equals(BigDecimal.class)) {
            result = NumberUtils.toScaledBigDecimal(cellValueString);
        } else if (fieldType.equals(Character.class)) {
            result = CharUtils.toChar(cellValueString);
        } else {
            result = cellValue;
        }
        return result;
    }

    private static Object getCellValue(Cell cell) {
        Object value;
        switch (cell.getCellTypeEnum()) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = NumberToTextConverter.toText(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            default:
                throw new RuntimeException(String.format("%s cell value is invalid, value is %s.",
                        cell.getAddress().formatAsString(), cell.getStringCellValue()));
        }
        return value;
    }

    /**
     * Get all fields of the class, the key is the name of the ExcelColumn annotation.
     *
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> HashMap<String, Field> getClassColumnNameMap(Class<T> clazz) {
        // Key is ExcelColumn annotation name
        HashMap<String, Field> columnNameMap = Maps.newHashMap();
        ArrayList<ExcelColumnAnnotationInfo> list = getClassExcelColumnAnnotationInfo(clazz);
        if (CollectionUtils.isEmpty(list)) {
            return columnNameMap;
        }
        for (ExcelColumnAnnotationInfo annotationInfo : list) {
            String value = annotationInfo.getName();
            if (columnNameMap.containsKey(value)) {
                throw new RuntimeException(String.format("%s class @%s annotation has the same name, name is the %s.",
                        clazz.getSimpleName(), ExcelColumn.class.getSimpleName(), value));
            }
            columnNameMap.put(value, annotationInfo.getField());
        }
        return columnNameMap;
    }

    /**
     * Get all fields of the class, the key is the number of the ExcelColumn annotation.
     *
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> TreeMap<Integer, ExcelColumnAnnotationInfo> getColumnNumberMap(Class<T> clazz) {
        // Key is ExcelColumn annotation number
        TreeMap<Integer, ExcelColumnAnnotationInfo> columnNumberMap = Maps.newTreeMap();
        ArrayList<ExcelColumnAnnotationInfo> list = getClassExcelColumnAnnotationInfo(clazz);
        if (CollectionUtils.isEmpty(list)) {
            return columnNumberMap;
        }
        for (ExcelColumnAnnotationInfo annotationInfo : list) {
            Integer number = annotationInfo.getNumber();
            if (columnNumberMap.containsKey(number)) {
                throw new RuntimeException(String.format("%s class @%s annotation has the same number, number is the %s.",
                        clazz.getSimpleName(), ExcelColumn.class.getSimpleName(), number));
            }
            columnNumberMap.put(number, annotationInfo);
        }
        return columnNumberMap;
    }

    private static <T> ArrayList<ExcelColumnAnnotationInfo> getClassExcelColumnAnnotationInfo(Class<T> clazz) {
        ArrayList<ExcelColumnAnnotationInfo> list = Lists.newArrayList();
        List<Field> fieldList = FieldUtils.getAllFieldsList(clazz);
        if (CollectionUtils.isEmpty(fieldList)) {
            return list;
        }
        Class<ExcelColumn> excelColumnClass = ExcelColumn.class;
        for (Field field : fieldList) {
            ExcelColumn annotation = field.getAnnotation(excelColumnClass);
            if (annotation == null) {
                continue;
            }
            list.add(new ExcelColumnAnnotationInfo(annotation.name(), annotation.number(), annotation.format(), field));
        }
        return list;
    }

    /**
     * Create header.
     *
     * @param sheet
     * @param clazz
     * @param <T>
     */
    private static <T> void createHeader(Sheet sheet, Class<T> clazz) {
        // Key is ExcelColumn annotation number
        TreeMap<Integer, ExcelColumnAnnotationInfo> columnNumberMap = getColumnNumberMap(clazz);
        Row row = sheet.createRow(sheet.getLastRowNum());
        Iterator<Map.Entry<Integer, ExcelColumnAnnotationInfo>> iterator = columnNumberMap.entrySet().iterator();
        while (iterator.hasNext()) {
            Map.Entry<Integer, ExcelColumnAnnotationInfo> next = iterator.next();
            Cell cell = row.createCell(next.getKey());
            cell.setCellValue(next.getValue().getName());
        }
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    static
    class ExcelColumnAnnotationInfo {

        /**
         * Column name
         */
        private String name;

        /**
         * Column number
         */
        private Integer number;

        /**
         * Column data format
         */
        private String format;

        /**
         * Class field
         */
        private Field field;
    }

    public enum Type {

        MICROSOFT_EXCEL(".xls"),
        OFFICE_OPEN_XML_SHEET(".xlsx");

        private String extensionName;

        Type(String extensionName) {
            this.extensionName = extensionName;
        }

        public String getExtensionName() {
            return extensionName;
        }
    }
}
