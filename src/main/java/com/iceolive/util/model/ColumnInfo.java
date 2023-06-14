package com.iceolive.util.model;

import com.iceolive.util.enums.ColumnType;
import lombok.Data;

import java.util.List;

/**
 * 列配置，用于导出Map
 *
 * @author wangmianzhe
 */
@Data
public class ColumnInfo {
    public  ColumnInfo(){

    }
    public ColumnInfo(String name,String title,ColumnType type){
        this.name = name;
        this.title = title;
        this.type = type;
    }
    /**
     * 字段名
     */
    private String name;
    /**
     * excel的标题
     */
    private String title;
    /**
     * 字段类型
     */
    private ColumnType type;
    /**
     * 校验规则
     */
    private List<Rule> rules;

    @Data
    public static class Rule {
        public Rule(){

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
