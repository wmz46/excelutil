package com.iceolive.util.enums;

import java.util.Arrays;

/**
 * @author wangmianzhe
 */

public enum ColumnType {
    /**
     * 字符串
     */
    STRING(0),

    /**
     * 整型
     */
    LONG(1),
    /**
     * 浮点型
     */
    DOUBLE(2),
    /**
     * 日期
     */
    DATE(3),
    /**
     * 日期时间
     */
    DATETIME(4),
    /**
     * 单张图片
     */
    IMAGE(5),
    /**
     * 多张图片
     */
    IMAGES(6);
    private int value;

    ColumnType(int value) {
        this.value = value;
    }

    public int getValue() {
        return this.value;
    }

    public static ColumnType valueOf(int value) {
        return Arrays.stream(values()).filter(m -> m.value == value).findFirst().orElse(null);
    }
}
