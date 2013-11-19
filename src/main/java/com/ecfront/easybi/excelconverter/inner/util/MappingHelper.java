package com.ecfront.easybi.excelconverter.inner.util;

import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 映射解析 辅助类
 */
public class MappingHelper {

    /**
     * 获取Bean中的自定义（非jre中内置）的子类
     * 警告：如果自定义类的包以java开头则需要使用Mapping(clazz=XX.class)来指定。
     */
    public static Class<?> getSubClass(Field field) {
        if (field.getGenericType() instanceof ParameterizedType) {
            //查找泛型容器的泛型class
            for (Type aType : ((ParameterizedType) field.getGenericType()).getActualTypeArguments()) {
                if (aType instanceof Class) {
                    return (Class) aType;
                }
            }
        } else {
            if (!field.getAnnotation(Map2J.class).clazz().equals(Object.class)) {
                return field.getAnnotation(Map2J.class).clazz();
            }
            if (!field.getType().isPrimitive() && !field.getType().getName().startsWith("java")) {
                return field.getType();
            }
        }
        return null;
    }

    /**
     * 解析单元格的行与列
     *
     * @param area  字符串格式的单元格
     * @param sheet 工作表
     * @return 行与列
     */
    public static int[] getCellIdx(String area, Sheet sheet) {
        int[] idx = {-1, -1};
        //匹配动态行
        boolean isFindDynamicRow = false;
        Matcher matcher = Pattern.compile("\\((-?[0-9]+)\\)").matcher(area);
        while (matcher.find()) {
            int dynamicRowIdx = Integer.valueOf(matcher.group(1));
            idx[0] = dynamicRowIdx >= 0 ? dynamicRowIdx : sheet.getLastRowNum() + dynamicRowIdx + 1;
            isFindDynamicRow = true;
            break;
        }
        if (!isFindDynamicRow) {
            //匹配静态行
            matcher = Pattern.compile("[^\\[\\(-]{1}([0-9]+)").matcher(area);
            while (matcher.find()) {
                idx[0] = Integer.valueOf(matcher.group(1)) - 1;
                break;
            }
        }
        //匹配动态列
        boolean isFindDynamicCol = false;
        matcher = Pattern.compile("\\[(-?[0-9]+)\\]").matcher(area);
        while (matcher.find()) {
            int dynamicColl = Integer.valueOf(matcher.group(1));
            idx[1] = dynamicColl >= 0 ? dynamicColl : idx[0] != -1 ? sheet.getRow(idx[0]).getLastCellNum() + dynamicColl + 1 : sheet.getRow(0).getLastCellNum() + dynamicColl + 1;
            isFindDynamicCol = true;
            break;
        }
        if (!isFindDynamicCol) {
            //匹配静态列
            matcher = Pattern.compile("[a-zA-Z]+").matcher(area);
            while (matcher.find()) {
                idx[1] = UnitConvertHelper.convert26to10(matcher.group()) - 1;
                break;
            }
        }
        return idx;
    }

}
