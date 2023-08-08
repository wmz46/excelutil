package com.iceolive.util.model;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Data
public class ImportSingleResult {
    private Map<String, Object> success;

    /**
     * 错误信息列表
     */
    private List<ImportResult.ErrorMessage> errors = new ArrayList<>();
}
