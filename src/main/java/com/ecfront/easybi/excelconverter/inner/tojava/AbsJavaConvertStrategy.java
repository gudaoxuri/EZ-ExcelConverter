package com.ecfront.easybi.excelconverter.inner.tojava;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;
import com.ecfront.easybi.excelconverter.inner.util.ExcelHelper;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

public abstract class AbsJavaConvertStrategy implements JavaConvertStrategy {

    //Bean Class
    protected Class beanClass;
    //工作表
    protected Sheet sheet;
    //合并过的单元格
    protected Map<int[], int[]> mergedCells;

    @Override
    public <E extends Object> E convert(Class<E> beanClass, Sheet sheet) throws Exception {
        this.beanClass = beanClass;
        this.sheet = sheet;
        mergedCells = ExcelHelper.findAllMergedCells(sheet);
        return (E) doConvert();
    }

    /**
     * 执行转换
     *
     * @throws Exception
     */
    protected abstract <E extends Object> E doConvert() throws Exception;

    /**
     * 解析、转换并返回最终值
     * Bean中要求的格式可与Excel各单元格的不一样，此方法承揽转换工作。
     *
     * @param field
     * @param rowIdx
     * @param colIdx
     * @return 最终值
     * @throws java.text.ParseException
     */
    protected Object getFinalValue(Field field, int rowIdx, int colIdx) throws ParseException {
        Class<?> fieldType = field.getType();
        Object val = ExcelHelper.getOriginalValue(rowIdx, colIdx, sheet);
        if (null == val || "".equals(val.toString())) {
            //忽略空值
            return null;
        }
        if (String.class.isAssignableFrom(fieldType)) {
            return val.toString();
        } else if (Long.class.isAssignableFrom(fieldType) || long.class.isAssignableFrom(fieldType)) {
            //舍掉小数取整
            return (long) Math.floor(Double.valueOf(val.toString()));
        } else if (Integer.class.isAssignableFrom(fieldType) || int.class.isAssignableFrom(fieldType)) {
            //舍掉小数取整
            return (int) Math.floor(Double.valueOf(val.toString()));
        } else if (Double.class.isAssignableFrom(fieldType) || double.class.isAssignableFrom(fieldType)) {
            return Double.valueOf(val.toString());
        } else if (Float.class.isAssignableFrom(fieldType) || float.class.isAssignableFrom(fieldType)) {
            return Float.valueOf(val.toString());
        } else if (Boolean.class.isAssignableFrom(fieldType) || boolean.class.isAssignableFrom(fieldType)) {
            return field.getAnnotation(Map2J.class).trueValue().equalsIgnoreCase(val.toString());
        } else if (Date.class.isAssignableFrom(fieldType)) {
            //优先使用excel中源格式，如果源格式是数值则进行转换，如果是字符串则按mapping中定义的格式转换
            if (val instanceof Date) {
                return val;
            } else if (val instanceof Double) {
                return new Date((long) Math.floor((Double) val));
            } else if (val instanceof String) {
                return new SimpleDateFormat(field.getAnnotation(Map2J.class).dateFormat()).parse((String) val);
            }
            return null;
        } else if (BigDecimal.class.isAssignableFrom(fieldType)) {
            return new BigDecimal(val.toString());
        } else {
            return val.toString();
        }
    }


}
