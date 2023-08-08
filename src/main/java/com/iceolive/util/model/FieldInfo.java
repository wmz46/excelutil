package com.iceolive.util.model;

import lombok.Data;


@Data
public class FieldInfo extends  BaseInfo{
    public FieldInfo(){

    }
    public FieldInfo(String name, String cellString, int type) {
        this.setName(name);
        this.cellString = cellString;
        this.setType(type);
    }

    /**
     * 单元格位置字符串
     */
    private String cellString;




}
