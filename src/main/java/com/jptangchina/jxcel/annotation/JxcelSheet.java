package com.jptangchina.jxcel.annotation;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface JxcelSheet {

    /**
     * sheet名称
     * @return sheet name
     */
    String value() default "";

    /**
     * 表头颜色，-1为不设置底色
     * 参考{@link org.apache.poi.ss.usermodel.IndexedColors}获取更多信息
     * @return color of head row
     */
    short color() default -1;

}
