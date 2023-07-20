package com.iceolive.util.enums;


/**
 * 规则类型
 * @author wangmianzhe
 */

public enum RuleType {
    /**
     * 内置方法
     */
    BUILTIN(0),
    /**
     * 正则
     */
    REGEXP(1),
    /**
     * 范围
     */
    RANGE(2),
    /**
     * 枚举
     */
    ENUMS(3);
    private Integer value;
    private RuleType(Integer value){
        this.value = value;
    }
    public Integer getValue(){
        return this.value;
    }
}
