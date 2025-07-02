package com.iceolive.util.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

/**
 * @author wmz
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelImportMapConfig {
    /**
     * Excel文件路径, 和inputStream二选一，inputStream优先于filepath
     */
    private String filepath;
    /**
     * Excel文件输入流,和filepath二选一，inputStream优先于filepath
     */
    private InputStream inputStream;
    /**
     * 列信息配置列表
     */
    private List<ColumnInfo> columnInfos;
    /**
     * 是否容错模式
     * true: 验证失败的数据跳过，继续处理后续数据
     * false: 遇到验证失败的数据立即停止处理
     */
    private boolean faultTolerant;
    /**
     * 开始行数，从1开始,默认1
     * 当onlyData = false时，包括标题行，且标题行只有一行，如当第一行是标题，则传1，当第二行是标题则传2。
     * 当onlyData = true时，不包括标题行，适合无标题行或多级标题行
     *
     */
    private int startRow = 1;
    /**
     * 自定义验证函数
     * 返回验证结果列表，空列表表示验证通过
     */
    private Function<Map<String, Object>, List<ValidateResult>> customValidateFunc;
    /**
     * 数据导入函数
     * 返回true表示导入成功，false或抛出异常表示导入失败
     */
    private Function<Map<String, Object>, Boolean> importFunc;

    /**
     * 开始行数是否只有数据，默认false
     * 当true时，适合无标题行或多级标题行
     */
    private boolean onlyData;


 }
