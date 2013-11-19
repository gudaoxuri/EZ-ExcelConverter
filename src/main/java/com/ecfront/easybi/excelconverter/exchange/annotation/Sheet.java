package com.ecfront.easybi.excelconverter.exchange.annotation;

import java.lang.annotation.*;

/**
 * 声明此bean是可用于转换的
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Sheet {

    /**
     * 指定sheet名称
     */
    String value() default "sheet1";

}
