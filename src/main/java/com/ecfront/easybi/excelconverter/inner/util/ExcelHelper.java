package com.ecfront.easybi.excelconverter.inner.util;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel 操作辅助类
 */
public class ExcelHelper {

    /**
     * 获取单元格中的值
     *
     * @param rowIdx 行号
     * @param colIdx 列号
     * @param sheet  工作表
     * @return 值
     */
    public static Object getOriginalValue(int rowIdx, int colIdx, Sheet sheet) {
        Cell cell = sheet.getRow(rowIdx).getCell(colIdx);
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue();
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    return cell.getRichStringCellValue();
                }
            case Cell.CELL_TYPE_ERROR:
                return null;
            default:
                return cell.getRichStringCellValue();
        }
    }

    /**
     * 获取单元格中的结构信息
     * // TODO  暂时只支持获取Drop down list信息
     *
     * @param rowIdx 行号
     * @param colIdx 列号
     * @param sheet  工作表
     * @return 值
     */
    public static List<String> getStruct(int rowIdx, int colIdx, Sheet sheet) {
        if (sheet instanceof XSSFSheet) {
            for (XSSFDataValidation validation : ((XSSFSheet) sheet).getDataValidations()) {
                for (CellRangeAddress cell : validation.getRegions().getCellRangeAddresses()) {
                    if (cell.isInRange(rowIdx, colIdx)) {
                        String[] result = validation.getValidationConstraint().getExplicitListValues();
                        //Bug array 首尾会多出"号
                        if (result.length > 0) {
                            result[0] = result[0].substring(1);
                            result[result.length - 1] = result[result.length - 1].substring(0, result[result.length - 1].length() - 1);
                        }
                        return Arrays.asList(result);
                    }
                }
            }
        }
        return null;
    }

    /**
     * 获取工作表中所有合并过的单元格
     *
     * @param sheet 工作表
     * @return 所有合并过的单元格，key为合并过的单元格，value为对应单元格所在合并区域的首个单元格，即有值的单元格
     */
    public static Map<int[], int[]> findAllMergedCells(Sheet sheet) {
        Map<int[], int[]> mergedCells = new HashMap<>();
        CellRangeAddress address;
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            address = sheet.getMergedRegion(i);
            for (int z = address.getFirstRow(); z <= address.getLastRow(); z++) {
                for (int y = address.getFirstColumn(); y <= address.getLastColumn(); y++) {
                    mergedCells.put(new int[]{z, y}, new int[]{address.getFirstRow(), address.getFirstColumn()});
                }
            }
        }
        return mergedCells;
    }

    /**
     * 获取合并过单元格有值的单元格，即所在合并区域的首个单元格
     *
     * @param cell        合并过的单元格
     * @param mergedCells 所有合并过的单元格
     * @return 所在合并区域的首个单元格
     */
    public static int[] getMergedValueCell(int[] cell, Map<int[], int[]> mergedCells) {
        for (int[] c : mergedCells.keySet()) {
            if (c[0] == cell[0] && c[1] == cell[1]) {
                return mergedCells.get(c);
            }
        }
        return cell;
    }

    public static CellRangeAddress getMergedRegion(Sheet sheet, Cell cell) {
        CellRangeAddress address;
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            address = sheet.getMergedRegion(i);
            if (address.getFirstRow() <= cell.getRowIndex() && address.getLastRow() >= cell.getRowIndex() && address.getFirstColumn() <= cell.getColumnIndex() && address.getLastColumn() >= cell.getColumnIndex()) {
                return address;
            }
        }
        return null;
    }

    public static Cell getCell(Sheet sheet, int rowIdx, int columnIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) {
            row = sheet.createRow(rowIdx);
        }
        return row.getCell(columnIdx) != null ? row.getCell(columnIdx) : row.createCell(columnIdx);
    }

    public static void setDataValidation(Sheet sheet, String[] textList, int firstRow, int endRow, int firstCol, int endCol) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        // 加载下拉列表内容
        DataValidationConstraint constraint = helper.createExplicitListConstraint(textList);
        constraint.setExplicitListValues(textList);
        // 设置数据有效性加载在哪个单元格上。
        // 四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList((short) firstRow, (short) endRow, (short) firstCol, (short) endCol);
        // 数据有效性对象
        DataValidation dataValidation = helper.createValidation(constraint, regions);
        sheet.addValidationData(dataValidation);
    }
}
