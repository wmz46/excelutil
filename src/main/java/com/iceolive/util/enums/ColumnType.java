package com.iceolive.util.enums;

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
     * 日期时间
     */
    DATETIME(3),
    /**
     * 单张图片
     */
    IMAGE(4),
    /**
     * 多张图片
     */
    IMAGES(5);
    private int value;

    ColumnType(int value) {
        this.value = value;
    }

    public int getValue() {
        return this.value;
    }
}
