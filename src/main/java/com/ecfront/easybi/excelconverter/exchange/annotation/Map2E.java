package com.ecfront.easybi.excelconverter.exchange.annotation;

import java.lang.annotation.*;


/**
 * 标记映射规则
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Map2E {

   String value();

    Mode mode() default Mode.ROW;

    String left() default "";

    String top() default "";

    String abs() default "";

    boolean isMatrix() default false;

    int colspan() default 0;

    String colspanExpression() default "";

    int rowspan() default 0;

    String rowspanExpression() default "";

}
