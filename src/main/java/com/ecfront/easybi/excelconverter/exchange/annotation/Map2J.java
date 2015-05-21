package com.ecfront.easybi.excelconverter.exchange.annotation;

import java.lang.annotation.*;


/**
 * 标记映射规则
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Map2J {

    /**
     * <h2>Excel中的单元格</h2>
     * 可用格式
     * <ul>
     * <li>某个单元格，如A2</li>
     * <li>某区域的单元格，用冒号分隔，如A2:B4</li>
     * <li>某几个单元格，用逗号分隔，如A2,B4</li>
     * <li>某一列，只用列号，如A</li>
     * <li>某区域列，只用列号，用冒号分隔，如A:B</li>
     * <li>某一行，只用行号，如1</li>
     * <li>某区域行，只用行号，用冒号分隔，如0:3</li></ul>
     * <h3>动态单元格</h3>
     * <p>定义及场景：有很多情况下单元格不是固定的，比如<strong>合计</strong>单元格就有可能随着项目列表（行）的增加而变化。</p>
     * <p>使用方法：行使用()包裹，列使用[]包裹，之中的数值是与行/列首、末的差值，格式与python数据一致，正数（0开始）表示与行/列首的差，负数（-1开始）表示与行/列末的差。</p>
     * 示例：假定表格区域在A0到B5
     * <ul>
     * <li>A(0)相当于A1，表示行首=1</li>
     * <li>B(2)相当于B3，表示行首加2行=1+2=3</li>
     * <li>B(-1)相当于B5，表示行末=5</li>
     * <li>B(-2)相当于B4，表示行末减1行=5-1=4</li>
     * <li>[2]0相当于C0，表示列首加2列=A+2=C</li>
     * <li>[2](2)相当于C3，表示列首加2列，行首加2行</li> </ul>
     * <p>警告：对于无法计算行号（只设置动态列，行号未指定且无法计算，计算方法参见下）的动态单元格会以工作表的第一行的最大列值做为此单元格的列尾值！</p>
     * <h3>父子区域设定</h3>
     * <p>定义及场景：在Excel映射Java bean时bean可能会由多个bean组合形成父子bean，此时子bean中所有的value有效范围都是父bean中对应成员变量定义范围的子集。</p>
     * <p> 示例：父类(Parent)中有一成员变量(List children)定义了范围为Mapping("A0:B2")，子类(Child)有一成员变量(name)定义了某一行Mapping("A")，在转换时这一变量(name)最终会被解析为A0:A2。</p>
     */
    String value();


    /**
     * 显式指定成员变量的class，如成员变量为泛型容器（List vals）或明确的类（SomeClass val）时系统可以自行解析，无需指定此值。
     */
    Class<?> clazz() default Object.class;

    /**
     * 当对应单元格要解析成boolean时用于定义为true的值，默认为"true"。
     */
    String trueValue() default "true";

    /**
     * 当对应单元格要解析成日期/时间，但Excel单元格却是字符串时用于定义转换格式，默认为"yyyy-MM-dd"。
     */
    String dateFormat() default "yyyy-MM-dd";

    /**
     * 是否获取结构信息，默认为false表示只获取值，为true用于获取结构信息，如下拉选项列表
     */
    boolean getStruct() default false;
}
