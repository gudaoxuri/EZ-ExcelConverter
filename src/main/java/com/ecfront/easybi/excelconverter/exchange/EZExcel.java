package com.ecfront.easybi.excelconverter.exchange;

import com.ecfront.easybi.excelconverter.inner.toexcel.ExcelAssembler;
import com.ecfront.easybi.excelconverter.inner.tojava.FastJavaConvertStrategy;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

/**
 * 转换入口
 */
public class EZExcel {

    /**
     * 执行转换
     *
     * @param filePath  Excel绝对路径
     * @param beanClass 目标Bean
     * @param <E>       Bean Class
     * @return 转换后的对象
     */
    public static <E extends Object> E toJava(String filePath, Class<E> beanClass) throws Exception {
        return toJava(new File(filePath), beanClass);
    }

    /**
     * 执行转换
     *
     * @param file      Excel文件
     * @param beanClass 目标Bean
     * @param <E>       Bean Class
     * @return 转换后的对象
     */
    public static <E extends Object> E toJava(File file, Class<E> beanClass) throws Exception {
        Workbook workbook;
        //选择格式
        if (file.getName().lastIndexOf("xlsx") != -1) {
            workbook = new XSSFWorkbook(new FileInputStream(file));
        } else {
            workbook = new HSSFWorkbook(new FileInputStream(file));
        }
        return toJava(workbook, beanClass);
    }

    /**
     * 执行转换
     *
     * @param workbook  workbook
     * @param beanClass 目标Bean
     * @param <E>       Bean Class
     * @return 转换后的对象
     */
    public static <E extends Object> E toJava(Workbook workbook, Class<E> beanClass) throws Exception {
        if (null != workbook && workbook.getNumberOfSheets() > 0) {
            if (beanClass.isAnnotationPresent(com.ecfront.easybi.excelconverter.exchange.annotation.Sheet.class)) {
                return doToJava(workbook, beanClass);
            }
        }
        return null;
    }

    private static <E extends Object> E doToJava(Workbook workbook, Class<E> beanClass) throws Exception {
        com.ecfront.easybi.excelconverter.exchange.annotation.Sheet annotationSheet = beanClass.getAnnotation(com.ecfront.easybi.excelconverter.exchange.annotation.Sheet.class);
        for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
            Sheet sheet = workbook.getSheetAt(numSheet);
            //查找匹配的sheet
            if (null != sheet &&
                    null != sheet.getSheetName() &&
                    sheet.getSheetName().equalsIgnoreCase(annotationSheet.value())) {
                //使用策略转换
                return new FastJavaConvertStrategy().convert(beanClass, sheet);
            }
        }
        return null;
    }

    public static File toExcel(Object bean, String filePath) throws Exception {
        File file = new File(filePath);
        new ExcelAssembler().assemble(bean, file, null);
        return file;
    }

    public static File toExcel(Object bean, String filePath, String templatePath) throws Exception {
        File file = new File(filePath);
        new ExcelAssembler().assemble(bean, file, new File(templatePath));
        return file;
    }

}
