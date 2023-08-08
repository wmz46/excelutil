package com.iceolive.util.model;

import com.iceolive.util.enums.RuleType;
import lombok.Data;

import java.util.List;

/**
 * @author wangmianzhe
 */
@Data
public class BaseInfo {
    /**
     * 字段名
     */
    private String name;
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

        public static Rule fromBuiltIn(String code) {
            Rule rule = new Rule();
            rule.type = RuleType.BUILTIN;
            rule.code = code;
            return rule;
        }

        /**
         * n
         *
         * @param code
         * @param message
         * @return
         */
        public static Rule fromBuiltIn(String code, String message) {
            Rule rule = new Rule();
            rule.type = RuleType.BUILTIN;
            rule.code = code;
            rule.message = message;
            return rule;
        }

        public static Rule fromRegExp(String code, String message) {
            Rule rule = new Rule();
            rule.type = RuleType.REGEXP;
            rule.code = code;
            rule.message = message;
            return rule;
        }

        public static Rule fromRange(Object min, Object max, String message) {
            Rule rule = new Rule();
            rule.type = RuleType.RANGE;
            rule.min = min;
            rule.max = max;
            rule.message = message;
            return rule;
        }

        public static Rule fromEnums(List<String> enumValues, String message) {
            Rule rule = new Rule();
            rule.type = RuleType.ENUMS;
            rule.enumValues = enumValues;
            rule.message = message;
            return rule;
        }

        /**
         * 校验规则，校验常量或 正则表达式
         */
        private String code;
        /**
         * 最小值，范围用
         */
        private Object min;
        /**
         * 最大值，范围用
         */
        private Object max;

        /**
         * 枚举值
         */
        private List<String> enumValues;
        /**
         * 错误提示语
         */
        private String message;
        private RuleType type;


    }
}
