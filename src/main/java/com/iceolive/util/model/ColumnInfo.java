package com.iceolive.util.model;

import com.iceolive.util.enums.ColumnType;
import lombok.Data;

import java.util.List;

/**
 * 列配置，用于导出Map
 * 优先使用colString识别
 * 如果没有才使用title识别
 *
 * @author wangmianzhe
 */
@Data
public class ColumnInfo {
    public ColumnInfo() {

    }

    /**
     * 列信息构造函数
     * @param name 列名
     * @param title excel的标题，如果没有列字母标识，则必填，否则非必填
     * @param colString 列标识
     * @param type 字段类型
     */

    public ColumnInfo(String name, String title, String colString, int type) {
        this.name = name;
        this.colString = colString;
        this.title = title;
        this.type = type;
    }

    /**
     * 字段名
     */
    private String name;
    /**
     * 列字母标识,建议必填
     */
    private String colString;
    /**
     * excel的标题，如果没有列字母标识，则必填，否则非必填
     */
    private String title;
    /**
     * 字段类型 ，对应枚举值 ColumnType
     */
    private int type;
    /**
     * 校验规则
     */
    private List<Rule> rules;


    @Data
    public static class Rule {
        public Rule() {

        }

        public Rule(String code, String message) {
            this.code = code;
            this.message = message;
        }

        /**
         * 校验规则，校验常量或 正则表达式
         * 正则表达式必须"/"开头和"/"结尾
         */
        private String code;
        /**
         * 错误提示语
         */
        private String message;

        public Rule(String code) {
            this.code = code;
        }


    }

}
