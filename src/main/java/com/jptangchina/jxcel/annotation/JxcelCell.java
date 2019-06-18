package com.jptangchina.jxcel.annotation;

import java.lang.annotation.*;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface JxcelCell {

    /**
     * 列名
     * @return column name
     */
    String value() default "";

    /**
     * 控制生成的表格中列的显示顺序
     * 默认是按照类中字段的定义顺序
     * @return order of head column
     */
    int order() default Short.MAX_VALUE;

    /**
     * 将数字值或布尔值与字符值转换，在导入导入都会生效
     * 例如：0 - 男，1 - 女
     * 如果字段为数字类型，那么值不能大于parse()数组的长度
     * 如果字段位布尔类型，应始终以0表示false，1表示true
     * @return semantic name
     */
    String[] parse() default {};

    /**
     * 如果字段为日期格式，可以在这里定义导入导出的日期格式
     * @return format of date
     */
    String format() default "yyyy-MM-dd HH:mm:ss:SSS";

    /**
     * 为字段添加后缀。
     * 例如可以为手机号添加"\t"后缀防止excel将字段识别为数字
     * @return suffix to add
     */
    String suffix() default "";

    /**
     * 列宽自适应
     * @return ture for autoresize
     */
    boolean autoResize() default true;
}
