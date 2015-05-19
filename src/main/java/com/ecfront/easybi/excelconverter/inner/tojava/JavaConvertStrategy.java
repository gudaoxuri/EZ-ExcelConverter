package com.ecfront.easybi.excelconverter.inner.tojava;


import org.apache.poi.ss.usermodel.Sheet;

/**
 * 转换策略
 */
public interface JavaConvertStrategy {

    /**
     * 转换
     *
     * @param beanClass 目标Bean
     * @param sheet     工作表
     * @param <E>       Bean Class
     * @return 转换后的对象
     */
    <E> E convert(Class<E> beanClass, Sheet sheet) throws Exception;
}
