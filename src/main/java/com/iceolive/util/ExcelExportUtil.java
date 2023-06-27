package com.iceolive.util;

import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.model.ColumnInfo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

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
            CreationHelper helper = workbook.getCreationHelper();
            Sheet sheet = workbook.getSheetAt(0);
            Drawing<?> drawing = sheet.getDrawingPatriarch();

            if(drawing== null){
                drawing = sheet.createDrawingPatriarch();
            }
            int r = startRow;
            if (!onlyData) {
                //填充标题
                Row row = sheet.getRow(r);
                if (row == null) {
                    row = sheet.createRow(r);
                }
                for (ColumnInfo columnInfo : columnInfos) {
                    if (StringUtil.isNotEmpty(columnInfo.getColString())) {
                        int c = CellReference.convertColStringToIndex(columnInfo.getColString());
                        Cell cell = row.getCell(c);
                        if (cell == null) {
                            cell = row.createCell(c);
                        }
                        cell.setCellValue(columnInfo.getTitle());
                    }
                }
                r++;
            }
            //填充数据
            for (Map<String, Object> item : data) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    row = sheet.createRow(r);
                }
                for (ColumnInfo columnInfo : columnInfos) {
                    if (StringUtil.isNotEmpty(columnInfo.getColString())) {
                        int c = CellReference.convertColStringToIndex(columnInfo.getColString());
                        Object value = item.get(columnInfo.getName());
                        Cell cell = row.getCell(c);
                        if (cell == null) {
                            cell = row.createCell(c);
                        }
                        switch (ColumnType.valueOf(columnInfo.getType())) {
                            case IMAGE:
                            case IMAGES:
                                if(value instanceof byte[]){
                                    ClientAnchor anchor1 = helper.createClientAnchor();
                                    anchor1.setCol1(c);
                                    anchor1.setRow1(r);
                                    anchor1.setCol2(c);
                                    anchor1.setRow2(r);
                                   loadPictureData(workbook, (byte[])value);
                                }else if(value.getClass().isAssignableFrom(ArrayList.class)){
                                    for (byte[] bytes : ((ArrayList<byte[]>) value)) {
                                        ClientAnchor anchor1 = helper.createClientAnchor();
                                        anchor1.setCol1(c);
                                        anchor1.setRow1(r);
                                        anchor1.setCol2(c);
                                        anchor1.setRow2(r);
                                        loadPictureData(workbook, bytes);
                                    }
                                }
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
                    }
                }

                r++;

            }
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    private static int loadPictureData(Workbook workbook, byte[] imageData)   {
        BufferedImage bufferedImage = ImageUtil.Bytes2Image(imageData);
        int pictureType;
        switch (bufferedImage.getType()){
            case BufferedImage.TYPE_INT_BGR:
                pictureType =  Workbook.PICTURE_TYPE_JPEG;
                break;
            case BufferedImage.TYPE_BYTE_GRAY:
            case BufferedImage.TYPE_3BYTE_BGR:
            case BufferedImage.TYPE_INT_ARGB:
            default:
                pictureType =Workbook.PICTURE_TYPE_PNG;
                break;
        }
        int pictureIndex = workbook.addPicture(imageData, pictureType);
        return pictureIndex;
    }

}
