package com.iceolive.util.annotation;

import javax.validation.constraints.Null;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author wangmianzhe
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)

public @interface ExcelColumn {
    //excel第一行标题名称
    String value() default "";


    String trueString() default "true";

    String falseString() default "false";

    //列宽，字符数 像素点 = 字符数*256
    int width() default -1;

    // 排序，导出用，越小排越前,默认100
    int order() default 100;
}
