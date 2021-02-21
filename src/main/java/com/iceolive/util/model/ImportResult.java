package com.iceolive.util.model;

import lombok.Data;
import org.apache.poi.ss.formula.functions.Index;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author wangmianzhe
 */
@Data
public class ImportResult<T> {

    /**
     * 成功的条数
     */
    private Map<Integer, T> successes = new LinkedHashMap<>();

    /**
     * 总条数
     */
    private int totalCount;


    /**
     * 获取成功的数据
     *
     * @return
     */
    public List<T> getSuccessList() {
        if (successes != null) {
            return successes.values().stream().collect(Collectors.toList());
        } else {
            return new ArrayList<>();
        }
    }

    /**
     * 获取成功的数量
     *
     * @return
     */
    public int getSuccessCount() {
        if (successes != null) {
            return successes.size();
        }
        return 0;
    }

    /**
     * 错误信息列表
     */
    private List<ErrorMessage> errors = new ArrayList<>();

    @Data
    public static class ErrorMessage {
        private Integer row;
        /**
         * excel的单元格地址 A1 B2
         */
        private String cell;

        /**
         * 错误信息
         */
        private String message;
    }
}
