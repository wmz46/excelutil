package com.iceolive.util;

import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.model.ColumnInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * @author wangmianzhe
 */
public class ExcelExportUtil {

    /**
     * 导出excel
     *
     * @param inputStream 导出模板
     * @param data        导出数据
     * @param columnInfos 导出列配置
     * @param startRow    导出数据开始行（从0开始）
     * @param onlyData    是否只导出数据（不含标题）
     * @return
     */
    public static byte[] exportExcel(
            InputStream inputStream,
            List<Map<String, Object>> data,
            List<ColumnInfo> columnInfos,
            int startRow,
            boolean onlyData
    ) {
        try {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int r = startRow;
            if (!onlyData) {
                //填充标题
                Row row = sheet.getRow(r);
                columnInfos.stream().filter(m -> StringUtil.isNotEmpty(m.getColString())).forEach(columnInfo ->
                {
                    int c = CellReference.convertColStringToIndex(columnInfo.getColString());
                    row.getCell(c).setCellValue(columnInfo.getTitle());
                });
                r++;
            }
            //填充数据
            for (Map<String, Object> item : data) {
                Row row = sheet.getRow(r);
                columnInfos.stream().filter(m -> StringUtil.isNotEmpty(m.getColString())).forEach(columnInfo ->
                {
                    int c = CellReference.convertColStringToIndex(columnInfo.getColString());
                    Object value = item.get(columnInfo.getName());
                    Cell cell = row.getCell(c);
                    switch (ColumnType.valueOf(columnInfo.getType())) {
                        case IMAGE:
                        case IMAGES:
                            break;
                        case LONG:
                            cell.setCellValue(Long.valueOf(String.valueOf(value)));
                            break;
                        case DOUBLE:
                            cell.setCellValue(Double.valueOf(String.valueOf(value)));
                            break;
                        case DATE:
                            cell.setCellValue(StringUtil.format(value, "yyyy-MM-dd"));
                            break;
                        case DATETIME:
                            cell.setCellValue(StringUtil.format(value, "yyyy-MM-dd HH:mm:ss"));
                            break;
                        case STRING:
                        default:
                            cell.setCellValue(String.valueOf(value));
                            break;
                    }
                });
                r++;
            }
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
