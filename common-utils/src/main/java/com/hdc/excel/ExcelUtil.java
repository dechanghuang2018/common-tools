package com.hdc.excel;

import com.alibaba.fastjson.JSON;
import com.hdc.excel.enums.ExcelTypeEnum;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * @Description: excel 工具类 只针对一个 sheet的情况
 * 功能：
 * 1. 获取真实的行数(默认第一行作为列名，所以不包括在内)
 * 2. 读取excel数据 解析成List<Map<String, Object>>，默认从第二行读取
 * 2. 读取excel数据 解析成List<T>，默认从第二行读取  格式要对，不然要出错
 *
 * @author hdc
 * @date 2019/3/18
 */
public class ExcelUtil {

    /**
     * 获取真实的行数（不包括第一行，若包括第一行，则 result + 1），不包括空行和只含空格的行
     *
     * @param inputStream
     * @param excelType
     * @param colNum      指定列的个数， 如果为空，则默认指定26列
     * @return
     * @throws Exception
     */
    public static int getRealNumOfRows(InputStream inputStream, ExcelTypeEnum excelType, Integer colNum) throws Exception {
        Workbook workbook = null;
        try {
            workbook = createWorkBook(inputStream, excelType);

            Sheet sheet = workbook.getSheetAt(0);
            // getLastRowNum: 没有数据或者只有第一行有数据返回0，最后有数据的行是第n行返回 n-1
            int lastRowNum = sheet.getLastRowNum();
            int realRowNum = 0;
            if (colNum == null) {
                //默认指定26列
                colNum = 26;
            }
            // 指定列数的单元格真实行数判断
            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                boolean isNullOfCell = isNullOfCell(row, colNum);
                if (row != null && !isNullOfCell)
                    realRowNum += 1;
            }
            return realRowNum;
        } catch (Exception e) {
            throw e;
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (Exception e) {
                // TODO
            }
        }
    }


    /**读取excel 解析成List<Map<String, Object>>，默认从第二行读取
     *
     * @param inputStream
     * @param excelType
     * @param columns  键，适用于 excel列名 == 实体类的属性 == 表字段
     * @param startRow
     * @return
     * @throws Exception
     */
    public static List<Map<String, Object>> readExcel(InputStream inputStream, ExcelTypeEnum excelType, String[] columns, int startRow) throws Exception {
        Workbook workbook = null;
        try {
            workbook = createWorkBook(inputStream, excelType);
            Sheet sheet = workbook.getSheetAt(0);
            int maxRow = sheet.getLastRowNum();
            int colNum = columns.length;
            // 解析数据
            List<Map<String, Object>> dataList = new ArrayList<>();
            for (int i = startRow; i <= maxRow; i++) {
                Row row = sheet.getRow(i);
                boolean isNullOfCell = isNullOfCell(row, colNum);
                if(row != null && !isNullOfCell){
                    Map<String, Object> map = new HashMap<>();
                    for (int col = 0; col < colNum; col++) {
                        map.put(columns[col], getCellValueOfCell(row.getCell(col)));
                    }
                    dataList.add(map);
                }
            }
            return dataList;
        } catch (Exception e) {
            throw  e;
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (Exception e) {
                // TODO
            }
        }
    }

    /**
     * 读取excel 解析成List<T>，默认从第二行读取  属性类型必须正确，不然报错。例如：时间格式
     *
     * @param inputStream
     * @param excelType
     * @param clazz
     * @param startRow
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> readExcel(InputStream inputStream, ExcelTypeEnum excelType, Class<T> clazz, int startRow) throws Exception {
        Field[] fields = clazz.getDeclaredFields();
        String[] modelNames = new String[fields.length];
        List<T> result = new ArrayList<>();
        for (int i = 0; i < fields.length; i++){
            // 获取属性的名字
            String name = fields[i].getName();
            modelNames[i] = name;
        }
        List<Map<String, Object>> dataList = readExcel(inputStream, excelType, modelNames, startRow);
        for (Map<String, Object> map: dataList){
            String str = JSON.toJSONString(map);
            result.add(JSON.parseObject(str, clazz));
        }
        return result;
    }

    public static Workbook createWorkBook(InputStream inputStream, ExcelTypeEnum excelType) throws Exception {
        Workbook workbook;
        if (ExcelTypeEnum.XLS.equals(excelType)) {
            workbook = (inputStream == null) ? new HSSFWorkbook() : new HSSFWorkbook(
                    new POIFSFileSystem(inputStream));
        } else {
            workbook = (inputStream == null) ? new SXSSFWorkbook(500) : new SXSSFWorkbook(
                    new XSSFWorkbook(inputStream));
        }
        return workbook;
    }

    /**
     * 获取cell的值
     *
     * @param cell
     * @return
     */
    public static Object getCellValueOfCell(Cell cell) {
        if (cell == null) {
            return null;
        }
        Object cellValue = null;
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType) {
            case STRING:
                cellValue = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                cellValue = getValueOfNumericCell(cell);
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                try {
                    cellValue = getValueOfNumericCell(cell);
                } catch (IllegalStateException e) {
                    try {
                        cellValue = cell.getRichStringCellValue().toString();
                    } catch (IllegalStateException e2) {
                        cellValue = cell.getErrorCellValue();
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            default:
                cellValue = cell.getRichStringCellValue().getString();
        }
        return cellValue;
    }

    /**
     * 获取数字型的cell值
     *
     * @param cell
     * @return
     */
    public static Object getValueOfNumericCell(Cell cell) {
        Boolean isDate = DateUtil.isCellDateFormatted(cell);
        Double d = cell.getNumericCellValue();
        Object obj = null;
        if (isDate) {
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            obj = dateFormat.format(cell.getDateCellValue());
        } else {
            obj = getRealStringValueOfDouble(d);
        }
        return obj;
    }

    /**
     * 处理科学计数法与普通计数法的字符串显示
     *
     * @param d
     * @return
     */
    public static String getRealStringValueOfDouble(Double d) {
        String value = d.toString();
        boolean b = value.contains("E");
        int indexOfPoint = value.indexOf('.');
        if (b) {
            int indexOfE = value.indexOf('E');
            // 小数部分
            BigInteger xs = new BigInteger(value.substring(indexOfPoint
                    + BigInteger.ONE.intValue(), indexOfE));
            // 指数
            int pow = Integer.valueOf(value.substring(indexOfE
                    + BigInteger.ONE.intValue()));
            int xsLen = xs.toByteArray().length;
            int scale = xsLen - pow > 0 ? xsLen - pow : 0;
            value = String.format("%." + scale + "f", d);
        } else {
            Pattern p = Pattern.compile(".0$");
            java.util.regex.Matcher m = p.matcher(value);
            if (m.find()) {
                value = value.replace(".0", "");
            }
        }
        return value;
    }

    /**
     * 指定列数的单元格非空判断，避免空行
     *
     * @param row
     * @param colNum
     * @return
     */
    public static boolean isNullOfCell(Row row, Integer colNum) {
        if (row == null) {
            return true;
        }
        for (int col = 0; col < colNum; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }
            // 获取单元格的值
            Object cellObj = getCellValueOfCell(cell);
            if (cellObj == null) {
                continue;
            }
            String cellValue = cellObj.toString();
            if (StringUtils.isNotBlank(cellValue)) {
                return false;
            }
        }
        return true;
    }

}
// 参考：https://www.cnblogs.com/yw0219/p/6625839.html
