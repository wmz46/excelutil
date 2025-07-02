package com.iceolive.util;

import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.enums.RuleType;
import com.iceolive.util.model.BaseInfo;
import com.iceolive.util.model.ColumnInfo;
import com.iceolive.util.model.ExcelExportMapConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.util.Units;

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
     * @param startRow    导出数据开始行（从1开始）
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
        return exportExcel(ExcelExportMapConfig.builder()
                .inputStream(inputStream)
                .data(data)
                .columnInfos(columnInfos)
                .startRow(startRow)
                .onlyData(onlyData)
                .build());
    }
    /**
     * 导出excel
     *
     * @param config 导出配置
     * @return
     */
    public static byte[] exportExcel(ExcelExportMapConfig config) {
        InputStream inputStream = config.getInputStream();
        List<Map<String, Object>> data = config.getData();
        int startRow = config.getStartRow();
        List<ColumnInfo> columnInfos = config.getColumnInfos();
        boolean onlyData = config.isOnlyData();
        int sheetIndex = config.getSheetIndex();
        int imgSize = 100;
        int imgPadding = 10;
        try {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            Drawing<?> drawing = sheet.getDrawingPatriarch();

            if (drawing == null) {
                drawing = sheet.createDrawingPatriarch();
            }
            int r = startRow - 1;
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
            int maxImageCount = 1;
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
                        if (value == null) {
                            continue;
                        }
                        switch (ColumnType.valueOf(columnInfo.getType())) {
                            case IMAGE:
                            case IMAGES:

                                //图片设置行高
                                row.setHeightInPoints((imgSize + 2 * imgPadding) * 0.75f);
                                float columnWidth = sheet.getColumnWidth(c);
                                int width = 32 * (imgSize + 2 * imgPadding);
                                if (columnWidth < width) {
                                    sheet.setColumnWidth(c, width);
                                }
                                if (value instanceof byte[]) {
                                    ClientAnchor anchor = new XSSFClientAnchor();
                                    anchor.setRow1(r);
                                    anchor.setCol1(c);
                                    anchor.setRow2(r);
                                    anchor.setCol2(c);
                                    anchor.setDx1(Units.EMU_PER_PIXEL * imgPadding);
                                    anchor.setDy1(Units.EMU_PER_PIXEL * imgPadding);
                                    anchor.setDx2(Units.EMU_PER_PIXEL * (imgSize + imgPadding));
                                    anchor.setDy2(Units.EMU_PER_PIXEL * (imgSize + imgPadding));
                                    anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
                                    Picture picture = drawing.createPicture(anchor, loadPictureData(workbook, (byte[]) value));

                                } else if (value.getClass().isAssignableFrom(ArrayList.class)) {
                                    ArrayList<byte[]> list = (ArrayList<byte[]>) value;
                                    int i = 0;
                                    for (byte[] bytes : list) {
                                        ClientAnchor anchor = new XSSFClientAnchor();
                                        anchor.setRow1(r);
                                        anchor.setCol1(c);
                                        anchor.setRow2(r);
                                        anchor.setCol2(c);
                                        anchor.setDx1(Units.EMU_PER_PIXEL * (i * (imgSize + imgPadding) + imgPadding));
                                        anchor.setDy1(Units.EMU_PER_PIXEL * imgPadding);
                                        anchor.setDx2(Units.EMU_PER_PIXEL * ((i + 1) * (imgSize + imgPadding)));
                                        anchor.setDy2(Units.EMU_PER_PIXEL * (imgSize + imgPadding));
                                        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
                                        Picture picture = drawing.createPicture(anchor, loadPictureData(workbook, bytes));

                                        i++;
                                    }

                                    if (!list.isEmpty()) {
                                        if (maxImageCount < list.size()) {
                                            maxImageCount = list.size();
                                        }
                                        width = 32 * (list.size() * (imgSize + imgPadding) + imgPadding);
                                        if (columnWidth < width) {
                                            sheet.setColumnWidth(c, width);
                                        }
                                    }

                                }
                                break;
                            case LONG:
                                cell.setCellValue(Long.parseLong(String.valueOf(value)));
                                break;
                            case DOUBLE:
                                cell.setCellValue(Double.parseDouble(String.valueOf(value)));
                                break;
                            case DATE:
                                cell.setCellValue(StringUtil.format(value, "yyyy-MM-dd"));
                                break;
                            case DATETIME:
                                cell.setCellValue(StringUtil.format(value, "yyyy-MM-dd HH:mm:ss"));
                                break;
                            case TIME:
                                cell.setCellValue(StringUtil.format(value, "HH:mm:ss"));
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
            byte[] bytes = baos.toByteArray();
            baos.close();
            return bytes;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static int loadPictureData(Workbook workbook, byte[] imageData) {
        BufferedImage bufferedImage = ImageUtil.Bytes2Image(imageData);
        int pictureType;
        switch (bufferedImage.getType()) {
            case BufferedImage.TYPE_INT_BGR:
                pictureType = Workbook.PICTURE_TYPE_JPEG;
                break;
            case BufferedImage.TYPE_BYTE_GRAY:
            case BufferedImage.TYPE_3BYTE_BGR:
            case BufferedImage.TYPE_INT_ARGB:
            default:
                pictureType = Workbook.PICTURE_TYPE_PNG;
                break;
        }
        return workbook.addPicture(imageData, pictureType);
    }

    private static double getScale(Picture picture) {
        // 最大宽度限制
        int maxWidth = 100;
        // 最大高度限制
        int maxHeight = 100;
        int originalWidth = picture.getImageDimension().width;
        int originalHeight = picture.getImageDimension().height;
        double scaleFactor = 1.0;
        if (originalWidth > maxWidth || originalHeight > maxHeight) {
            double widthRatio = (double) maxWidth / originalWidth;
            double heightRatio = (double) maxHeight / originalHeight;
            scaleFactor = Math.min(widthRatio, heightRatio);
        }
        return scaleFactor;
    }
    /**
     * 设置excel单元格有效性
     * @param inputStream excel模板流
     * @param columnInfos 列信息
     * @param startRow 开始行，从0开始
     * @return 返回新的excel字节数组
     */
    public static byte[] setDataValidationRules(
            InputStream inputStream,
            List<ColumnInfo> columnInfos,
            int startRow
    ) {
        return setDataValidationRules(ExcelExportMapConfig.builder()
                .inputStream(inputStream)
                .columnInfos(columnInfos)
                .startRow(startRow)
                .build());
    }
    /**
     * 设置excel单元格有效性
     * @param config 导出配置
     * @return 返回新的excel字节数组
     */
    public static byte[] setDataValidationRules( ExcelExportMapConfig config    ) {
        try {
            InputStream inputStream = config.getInputStream();
            List<ColumnInfo> columnInfos = config.getColumnInfos();
            int startRow = config.getStartRow();
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(config.getSheetIndex());
            int sheetIndex = workbook.getNumberOfSheets();
            String hiddenSheetName = "_hiddenSheet";
            Sheet hiddenSheet = workbook.createSheet(hiddenSheetName);
            workbook.setSheetHidden(sheetIndex,true);

            for (ColumnInfo columnInfo : columnInfos) {
                if (StringUtil.isNotEmpty(columnInfo.getColString())) {
                    int c = CellReference.convertColStringToIndex(columnInfo.getColString());
                    if (columnInfo.getRules() == null) {
                        continue;
                    }
                    for (BaseInfo.Rule rule : columnInfo.getRules().stream().filter(m -> m.getType() == RuleType.ENUMS || m.getType() == RuleType.RANGE).collect(Collectors.toList())) {
                        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
                        DataValidationConstraint dvConstraint;
                        if (rule.getType() == RuleType.RANGE) {
                            dvConstraint = dvHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.BETWEEN, rule.getMin().toString(), rule.getMax().toString());
                        } else {
                            for (int i = 0; i < rule.getEnumValues().size(); i++) {
                                String enumValue = rule.getEnumValues().get(i);
                                Row row = hiddenSheet.getRow(i);
                                if (row == null) {
                                    row = hiddenSheet.createRow(i);
                                }
                                Cell cell = row.getCell(c);
                                if (cell == null) {
                                    cell = row.createCell(c);
                                }
                                cell.setCellValue(enumValue);
                            }
                            String strFormula = hiddenSheetName + "!$" + columnInfo.getColString() + "$1:$" + columnInfo.getColString() + "$" + rule.getEnumValues().size();
                            dvConstraint = new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.LIST, strFormula);
                        }
                        // 设置数据有效性在哪个区域
                        CellRangeAddressList addressList = new CellRangeAddressList(startRow, 65535, c , c );
                        // 创建数据有效性对象
                        DataValidation dataValidation = dvHelper.createValidation(dvConstraint, addressList);

                        dataValidation.setShowErrorBox(true);
                        if (StringUtil.isBlank(rule.getMessage())) {
                            dataValidation.createErrorBox("错误", "您输入的值不是有效值");
                        } else {
                            dataValidation.createErrorBox("错误", rule.getMessage());
                        }
                        // 将数据有效性应用到工作表
                        sheet.addValidationData(dataValidation);
                    }
                }
            }
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            byte[] bytes = baos.toByteArray();
            baos.close();
            return bytes;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
