package com.ecfront.easybi.excelconverter.exchange.annotation;

import java.lang.annotation.*;


/**
 * 标记映射规则
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Validation {

     String left() default "";

    String top() default "";

    String abs() default "";

     String length() default "";

    Mode mode() default Mode.ROW;

    int absLength() default -1;

}
