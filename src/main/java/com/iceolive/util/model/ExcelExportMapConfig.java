package com.iceolive.util.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * @author wmz
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelExportMapConfig {
    private InputStream inputStream;
    private List<Map<String, Object>> data;
    private List<ColumnInfo> columnInfos;
    private int startRow = 1;
    boolean onlyData = false;
    private int sheetIndex = 0;
}
