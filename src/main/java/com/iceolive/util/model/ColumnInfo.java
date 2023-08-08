package com.iceolive.util.model;

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
public class ColumnInfo extends BaseInfo {
    public ColumnInfo() {

    }

    /**
     * 列信息构造函数
     *
     * @param name      列名
     * @param title     excel的标题，如果没有列字母标识，则必填，否则非必填
     * @param colString 列标识
     * @param type      字段类型
     */

    public ColumnInfo(String name, String title, String colString, int type) {
        this.setName(name);
        this.colString = colString;
        this.title = title;
        this.setType(type);
    }


    /**
     * 列字母标识,建议必填
     */
    private String colString;
    /**
     * excel的标题，如果没有列字母标识，则必填，否则非必填
     */
    private String title;

    /**
     * 校验规则
     */
    private List<Rule> rules;




}
