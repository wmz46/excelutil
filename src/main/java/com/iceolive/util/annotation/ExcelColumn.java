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
    /**
     * excel标题名称，有换行符需加换行符\n
     * @return
     */
    String value() default "";

    /**
     * 列标识  A B C D
     * 优先级大于value
     * 当设置onlyData时，此字段必填
     *
     * @return
     */
    String colString() default "";

    String trueString() default "true";

    String falseString() default "false";


}
