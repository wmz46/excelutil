package com.iceolive.util.annotation;

import java.lang.annotation.*;

/**
 * @author wangmianzhe
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {
    //导出名称
    String name();

    //行高，字符数 像素点 约等于 字符数*256
    int height() default -1;
}
