package com.ecfront.easybi.excelconverter.inner.tojava;

import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;
import com.ecfront.easybi.excelconverter.inner.util.ExcelHelper;
import com.ecfront.easybi.excelconverter.inner.util.MappingHelper;

import java.lang.reflect.Field;
import java.text.ParseException;
import java.util.*;

/**
 * 默认的转换策略
 * <ul>此策略处理步骤如下：
 * <li><ul>
 * <li>遍历成员变量，递归bean，获取所有被Mapping标记的成员变量</li>
 * <li>根据层级关系计算各Mapping的最终value</li></ul></li>
 * <li><ul><li>遍历成员变量，递归bean，结合对应的Mapping</li></ul></li>
 * </ul>
 */
public class FastJavaConvertStrategy extends AbsJavaConvertStrategy {


    private Map<String, List<int[]>> MAPPING = new HashMap<String, List<int[]>>();

    @Override
    public Object doConvert() throws Exception {
        assembleMapping(beanClass, null);
        return assembleBean(beanClass, 0);
    }

    //组装MAPPING，获取所有field和其对应的单元格
    private <O extends Object> void assembleMapping(Class<O> beanClass, List<int[]> parentArea) {
        String area;
        Class<?> subClass;
        for (Field field : beanClass.getDeclaredFields()) {
            if (!field.isAnnotationPresent(Map2J.class)) {
                continue;
            }
            area = field.getAnnotation(Map2J.class).value();
            if (null == area || "".equals(area)) {
                continue;
            }
            //保存字段与区域
            MAPPING.put(beanClass.getName() + "." + field.getName(), assembleArea(area, parentArea));
            subClass = MappingHelper.getSubClass(field);
            if (null != subClass) {
                //递归bean
                assembleMapping(subClass, MAPPING.get(beanClass.getName() + "." + field.getName()));
            }
        }
    }

    //组装区域
    private List<int[]> assembleArea(String area, List<int[]> parentArea) {
        List<int[]> resultAreas = new ArrayList<int[]>();
        if (-1 == area.indexOf(":") && -1 == area.indexOf(",")) {
            int[] idx = MappingHelper.getCellIdx(area, sheet);
            if (-1 == idx[0] || -1 == idx[1]) {
                //把类似"A"这样的区域改为"A:A"，因为它本身就是一个区域
                area = area + ":" + area;
            }
        }
        if (-1 != area.indexOf(":")) {
            //单元格区域处理
            //获取首末单元格，对没有指定相应行/列的（多在子对象中）按父区域的边界组装
            int[] startIdx = MappingHelper.getCellIdx(area.split(":")[0], sheet);
            int[] endIdx = MappingHelper.getCellIdx(area.split(":")[1], sheet);
            if (-1 == startIdx[0]) {
                startIdx[0] = getRowNum(parentArea, false);
            }
            if (-1 == startIdx[1]) {
                startIdx[1] = getColNum(parentArea, startIdx[0], false);
            }
            if (-1 == endIdx[0]) {
                endIdx[0] = getRowNum(parentArea, true);
            }
            if (-1 == endIdx[1]) {
                endIdx[1] = getColNum(parentArea, endIdx[0], true);
            }
            for (int i = startIdx[0]; i <= endIdx[0]; i++) {
                for (int j = startIdx[1]; j <= endIdx[1]; j++) {
                    addArea(new int[]{i, j}, parentArea, resultAreas);
                }
            }
        } else if (-1 != area.indexOf(",")) {
            //多个单元格处理
            String[] areas = area.split(",");
            for (String a : areas) {
                addArea(a, parentArea, resultAreas);
            }
        } else {
            //单一单元格处理
            addArea(area, parentArea, resultAreas);
        }
        return resultAreas;
    }

    //根据父区域或工作表可用区域获取行首/末
    private int getRowNum(List<int[]> parentArea, boolean isLast) {
        if (null == parentArea || parentArea.size() == 0) {
            if (isLast) {
                return sheet.getLastRowNum();
            } else {
                return sheet.getFirstRowNum();
            }
        } else {
            if (isLast) {
                return parentArea.get(parentArea.size() - 1)[0];
            } else {
                return parentArea.get(0)[0];
            }
        }
    }

    //根据父区域或工作表可用区域获取列首/末，父区域不可用时列末需要指定行号
    private int getColNum(List<int[]> parentArea, int rowIdx, boolean isLast) {
        if (null == parentArea || parentArea.size() == 0) {
            if (isLast) {
                return sheet.getRow(rowIdx).getLastCellNum();
            } else {
                return sheet.getRow(rowIdx).getFirstCellNum();
            }
        } else {
            if (isLast) {
                return parentArea.get(parentArea.size() - 1)[1];
            } else {
                return parentArea.get(0)[1];
            }
        }
    }

    private void addArea(String area, List<int[]> parentArea, List<int[]> resultAreas) {
        int[] idx = MappingHelper.getCellIdx(area, sheet);
        addArea(idx, parentArea, resultAreas);
    }

    private void addArea(int[] idx, List<int[]> parentArea, List<int[]> resultAreas) {
        idx = ExcelHelper.getMergedValueCell(idx, mergedCells);
        if (isAddAreaAble(idx, parentArea)) {
            resultAreas.add(idx);
        }
    }

    //是否可以添加，要求指定的单元格在父区域内
    private boolean isAddAreaAble(int[] idx, List<int[]> parentArea) {
        if (null == idx || idx[0] < 0 || idx[1] < 0) {
            return false;
        }
        if (null == parentArea || parentArea.size() == 0) {
            return true;
        }
        int len = parentArea.size();
        for (int i = 0; i < len; i++) {
            if (idx[0] == parentArea.get(i)[0] && idx[1] == parentArea.get(i)[1]) {
                return true;
            }
        }
        return false;
    }

    //组装Bean，mapIdx主要用于容器成员变量，每次循环加1
    private <O extends Object> O assembleBean(Class<O> beanClass, int mapIdx) throws IllegalAccessException, InstantiationException, ParseException {
        O obj = beanClass.newInstance();
        Class<?> subClass;
        Object result;
        for (Field field : beanClass.getDeclaredFields()) {
            if (!field.isAnnotationPresent(Map2J.class)) {
                continue;
            }
            field.setAccessible(true);
            subClass = MappingHelper.getSubClass(field);
            if (null != subClass) {
                if (Collection.class.isAssignableFrom(field.getType())) {
                    //容器组装，窗口的size由 MAPPING 确定
                    Collection collection = (Collection) field.get(obj);
                    //获取对应字段的区域列表
                    List<int[]> m = MAPPING.get(beanClass.getName() + "." + field.getName());
                    if (null != m && m.size() > 0) {
                        //根据区域中首末单元格行号的差计算size
                        int len = m.get(m.size() - 1)[0] - m.get(0)[0] + 1;
                        for (int i = 0; i < len; i++) {
                            collection.add(assembleBean(subClass, i));
                        }
                    }
                } else {
                    field.set(obj, assembleBean(subClass, mapIdx));
                }
            } else {
                result = getFieldValue(beanClass.getName() + "." + field.getName(), field, mapIdx);
                if (null != result) {
                    field.set(obj, result);
                }
            }
        }
        return obj;
    }

    //获取结果值
    private Object getFieldValue(String key, Field field, int mapIdx) throws ParseException {
        int[] area = MAPPING.get(key).get(mapIdx);
        return getFinalValue(field, area[0], area[1]);
    }

}
