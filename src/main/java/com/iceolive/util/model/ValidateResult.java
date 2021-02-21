package com.iceolive.util.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 用于验证失败存储的信息，框架会根据字段名自动找到单元格地址。
 * @author wangmianzhe
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ValidateResult {
    /**
     * 中间类字段名
     */
    private String fieldName;
    /**
     * 错误信息
     */
    private String message;
}