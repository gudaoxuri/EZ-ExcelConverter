package com.ecfront.easybi.excelconverter.inner.toexcel;

import com.ecfront.easybi.excelconverter.exchange.annotation.Map2E;
import com.ecfront.easybi.excelconverter.exchange.annotation.Mode;
import com.ecfront.easybi.excelconverter.exchange.annotation.Validation;
import com.ecfront.easybi.excelconverter.inner.util.ExcelHelper;
import com.ecfront.easybi.excelconverter.inner.util.MappingHelper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

public class ExcelAssembler {
    private Map<String, Field> orderDataItems = new LinkedHashMap<>();
    private Map<String, int[]> dataItemborders = new HashMap<>();
    private Object bean;
    private Sheet sheet;

    public void assemble(Object bean, File outputFile, File templateFile) throws Exception {
        this.bean = bean;
        com.ecfront.easybi.excelconverter.exchange.annotation.Sheet annotationSheet = bean.getClass().getAnnotation(com.ecfront.easybi.excelconverter.exchange.annotation.Sheet.class);
        Workbook workbook;
        if (null == templateFile) {
            workbook = new XSSFWorkbook();
            if (annotationSheet.value() != null && !"".equals(annotationSheet.value().trim())) {
                sheet = workbook.createSheet(annotationSheet.value());
            } else {
                sheet = workbook.createSheet();
            }
        } else {
            String extName = templateFile.getName().substring(templateFile.getName().lastIndexOf(".") + 1);
            if ("xlsx".equalsIgnoreCase(extName)) {
                workbook = new XSSFWorkbook(new FileInputStream(templateFile));
            } else if ("xls".equalsIgnoreCase(extName)) {
                workbook = new HSSFWorkbook(new FileInputStream(templateFile));
            } else {
                throw new Exception("Template File is NOT allowed!");
            }
            if (annotationSheet.value() != null && !"".equals(annotationSheet.value().trim())) {
                sheet = workbook.getSheet(annotationSheet.value());
            } else {
                sheet = workbook.getSheetAt(0);
            }
        }
        parseDependent(bean);
        packageDataItems();
        packageValidations();
        FileOutputStream os = new FileOutputStream(outputFile);
        workbook.write(os);
        os.close();
    }

    private void parseDependent(Object bean) {
        Field[] fields = bean.getClass().getFields();
        if (fields != null && fields.length > 0) {
            Map<String, Object[]> allDataItems = new HashMap<>();
            Map2E map;
            List<String> dependents;
            //1st get all data items
            for (Field f : fields) {
                if (f.isAnnotationPresent(Map2E.class)) {
                    map = f.getAnnotation(Map2E.class);
                    dependents = new ArrayList<>();
                    if (!"".equals(map.left())) {
                        dependents.addAll(Arrays.asList(map.left().split(",")));
                    }
                    if (!"".equals(map.top())) {
                        dependents.addAll(Arrays.asList(map.top().split(",")));
                    }
                    allDataItems.put(map.value(), new Object[]{dependents, f});
                }
            }
            if (allDataItems.size() > 0) {
                List<String> orderDataItemNames = new LinkedList<>();
                for (Map.Entry<String, Object[]> entity : allDataItems.entrySet()) {
                    addOrderDataItemName(orderDataItemNames, entity.getKey(), (List<String>) entity.getValue()[0], allDataItems);
                }
                for (String s : orderDataItemNames) {
                    orderDataItems.put(s, (Field) allDataItems.get(s)[1]);
                }
            }
        }
    }

    private void addOrderDataItemName(List<String> orderDataItemNames, String title, List<String> dependents, Map<String, Object[]> allDataItems) {
        if (orderDataItemNames.contains(title)) {
            return;
        }
        if (dependents.size() > 0) {
            //the dependent is not add to orderDataItemNames yet
            dependents.stream().filter(s -> orderDataItemNames.indexOf(s) == -1).forEach(s -> {
                //the dependent is not add to orderDataItemNames yet
                Object[] entity = allDataItems.get(s);
                addOrderDataItemName(orderDataItemNames, s, (List<String>) entity[0], allDataItems);
            });
            int maxIdx = 0;
            int currentIdx;
            for (String s : dependents) {
                currentIdx = orderDataItemNames.indexOf(s);
                if (currentIdx > maxIdx) {
                    maxIdx = currentIdx;
                }
            }
            orderDataItemNames.add(maxIdx + 1, title);
        } else {
            orderDataItemNames.add(0, title);
        }
    }

    private void packageDataItems() throws IllegalAccessException {
        if (orderDataItems.size() == 0) {
            return;
        }
        Map2E map;
        Cell cell;
        int[] lastIdxs = null;
        Object value;
        for (Map.Entry<String, Field> entity : orderDataItems.entrySet()) {
            map = entity.getValue().getAnnotation(Map2E.class);
            value = entity.getValue().get(bean);
            if (value != null) {
                cell = getFirstCellByMap2E(map);
                lastIdxs = new int[]{cell.getRowIndex(), cell.getColumnIndex()};
                packageDataItem(value, cell, map.mode(), map.isMatrix(), lastIdxs, -1, 0);
            }
            dataItemborders.put(entity.getKey(), lastIdxs);
        }
    }

    private Cell packageDataItem(Object value, Cell cell, Mode mode, boolean isMatrix, int[] lastIdxs, int lastIdx, int level) {
        if (value instanceof Collection || value instanceof Object[]) {
            Collection values;
            Iterator valuesIterator;
            if (value instanceof Object[]) {
                values = Arrays.asList((Object[]) value);
            } else {
                values = (Collection) value;
            }
            if (!isMatrix && lastIdx == -1) {
                lastIdx = detectDeep(values.iterator(), cell, mode);
            }
            valuesIterator = values.iterator();
            boolean rowMode;
            while (valuesIterator.hasNext()) {
                packageDataItem(valuesIterator.next(), cell, mode, isMatrix, lastIdxs, lastIdx, level + 1);
                if (valuesIterator.hasNext()) {
                    CellRangeAddress address = ExcelHelper.getMergedRegion(sheet, cell);
                    if (isMatrix) {
                        rowMode = level % 2 != 0;
                    } else {
                        rowMode = mode == Mode.ROW;
                    }
                    if (rowMode) {
                        if (address == null) {
                            cell = ExcelHelper.getCell(sheet, cell.getRowIndex(), cell.getColumnIndex() + 1);
                        } else {
                            cell = ExcelHelper.getCell(sheet, cell.getRowIndex(), address.getLastColumn() + 1);
                        }
                    } else {
                        if (address == null) {
                            cell = ExcelHelper.getCell(sheet, cell.getRowIndex() + 1, cell.getColumnIndex());
                        } else {
                            cell = ExcelHelper.getCell(sheet, address.getLastRow() + 1, cell.getColumnIndex());
                        }
                    }
                }
            }
        } else if (value instanceof Map) {
            Cell lastCell;
            for (Map.Entry<String, Object> entity : ((Map<String, Object>) value).entrySet()) {
                cell.setCellValue(entity.getKey());
                if (entity.getValue() != null) {
                    if (mode == Mode.COLUMN) {
                        lastCell = packageDataItem(entity.getValue(), ExcelHelper.getCell(sheet, cell.getRowIndex(), cell.getColumnIndex() + 1), mode, isMatrix, lastIdxs, lastIdx, level + 1);
                    } else {
                        lastCell = packageDataItem(entity.getValue(), ExcelHelper.getCell(sheet, cell.getRowIndex() + 1, cell.getColumnIndex()), mode, isMatrix, lastIdxs, lastIdx, level + 1);
                    }
                    CellRangeAddress address = ExcelHelper.getMergedRegion(sheet, lastCell);
                    if (mode == Mode.ROW) {
                        if (address == null) {
                            sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), lastCell.getColumnIndex()));
                        } else {
                            sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), address.getLastColumn()));
                        }
                    } else {
                        if (address == null) {
                            sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), lastCell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex()));
                        } else {
                            sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), address.getLastRow(), cell.getColumnIndex(), cell.getColumnIndex()));
                        }
                    }
                }
            }
        } else {
            if (value instanceof String) {
                cell.setCellValue((String) value);
            } else if (value instanceof Date) {
                cell.setCellValue((Date) value);
            } else if (value instanceof Integer) {
                cell.setCellValue((Integer) value);
            } else if (value instanceof Float) {
                cell.setCellValue((Float) value);
            } else if (value instanceof Double) {
                cell.setCellValue((Double) value);
            } else if (value instanceof BigDecimal) {
                cell.setCellValue(((BigDecimal) value).doubleValue());
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else if (value instanceof Calendar) {
                cell.setCellValue((Calendar) value);
            }
            if (!isMatrix && 0 != lastIdx) {
                if (mode == Mode.ROW) {
                    if (lastIdx > cell.getRowIndex()) {
                        sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), lastIdx, cell.getColumnIndex(), cell.getColumnIndex()));
                        if (lastIdxs[0] < lastIdx) {
                            lastIdxs[0] = lastIdx;
                        }
                    }
                } else {
                    if (lastIdx > cell.getColumnIndex()) {
                        sheet.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), lastIdx));
                        if (lastIdxs[1] < lastIdx) {
                            lastIdxs[1] = lastIdx;
                        }
                    }
                }
            }
        }
        if (lastIdxs[0] < cell.getRowIndex()) {
            lastIdxs[0] = cell.getRowIndex();
        }
        if (lastIdxs[1] < cell.getColumnIndex()) {
            lastIdxs[1] = cell.getColumnIndex();
        }
        return cell;
    }

    private void packageValidations() throws IllegalAccessException {
        Field[] fields = bean.getClass().getFields();
        if (fields != null && fields.length > 0) {
            Validation validation;
            Integer[] firstIdxs;
            int endRow;
            int endColumn;
            for (Field f : fields) {
                if (f.isAnnotationPresent(Validation.class)) {
                    validation = f.getAnnotation(Validation.class);
                    firstIdxs = getFirstIdx(validation.abs(), validation.top(), validation.left());
                    if (Mode.COLUMN == validation.mode()) {
                        endColumn = firstIdxs[1];
                        if (-1 != validation.absLength()) {
                            endRow = firstIdxs[0] + validation.absLength() - 1;
                        } else {
                            endRow = firstIdxs[0] + getValueLength(orderDataItems.get(validation.length()).get(bean)) - 1;
                        }
                    } else {
                        endRow = firstIdxs[0];
                        if (-1 != validation.absLength()) {
                            endColumn = firstIdxs[1] + validation.absLength() - 1;
                        } else {
                            endColumn = firstIdxs[1] + getValueLength(orderDataItems.get(validation.length()).get(bean)) - 1;
                        }
                    }
                    ExcelHelper.setDataValidation(sheet, (String[]) f.get(bean), firstIdxs[0], endRow, firstIdxs[1], endColumn);
                }
            }
        }
    }

    private int getValueLength(Object value) {
        if (value instanceof Collection) {
            return ((Collection) value).size();
        } else if (value instanceof Object[]) {
            return ((Object[]) value).length;
        } else if (value instanceof Map) {
            return ((Map) value).size();
        } else {
            return 1;
        }
    }

    private int _tmpMaxDeep;

    private int detectDeep(Iterator values, Cell cell, Mode mode) {
        _tmpMaxDeep = 0;
        detectDeep(values, 0);
        if (mode == Mode.ROW) {
            return _tmpMaxDeep + cell.getRowIndex();
        } else {
            return _tmpMaxDeep + cell.getColumnIndex();
        }
    }

    private void detectDeep(Iterator values, int currentDeep) {
        Object value;
        while (values.hasNext()) {
            value = values.next();
            if (value instanceof Map) {
                Collection subValues;
                for (Map.Entry<String, Object> entity : ((Map<String, Object>) value).entrySet()) {
                    if (entity.getValue() != null) {
                        if (entity.getValue() instanceof Object[]) {
                            subValues = Arrays.asList((Object[]) entity.getValue());
                        } else {
                            subValues = (Collection) entity.getValue();
                        }
                        detectDeep(subValues.iterator(), currentDeep + 1);
                    }
                }
            }
        }
        if (_tmpMaxDeep < currentDeep) {
            _tmpMaxDeep = currentDeep;
        }
    }

    private Cell getFirstCellByMap2E(Map2E map) {
        return getFirstCell(map.abs(), map.top(), map.left());
    }

    private Cell getFirstCellByValidation(Validation validation) {
        return getFirstCell(validation.abs(), validation.top(), validation.left());
    }

    private Cell getFirstCell(String tAbs, String tTop, String tLeft) {
        Integer[] idxs = getFirstIdx(tAbs, tTop, tLeft);
        return ExcelHelper.getCell(sheet, idxs[0], idxs[1]);
    }

    private Integer[] getFirstIdx(String tAbs, String tTop, String tLeft) {
        int left = 0;
        int top = 0;
        if (!"".equals(tAbs)) {
            int[] abs = MappingHelper.getCellIdx(tAbs, sheet);
            left = abs[1];
            top = abs[0];
        } else {
            if (!"".equals(tLeft)) {
                for (String s : tLeft.split(",")) {
                    if (left < dataItemborders.get(s)[1]) {
                        left = dataItemborders.get(s)[1];
                    }
                }
                left++;
            }
            if (!"".equals(tTop)) {
                for (String s : tTop.split(",")) {
                    if (top < dataItemborders.get(s)[0]) {
                        top = dataItemborders.get(s)[0];
                    }
                }
                top++;
            }
        }
        return new Integer[]{top, left};
    }
}
