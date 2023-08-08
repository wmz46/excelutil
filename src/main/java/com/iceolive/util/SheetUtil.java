package com.iceolive.util;

import com.iceolive.util.model.CellImages;
import com.iceolive.util.model.CellImagesRels;
import com.iceolive.xpathmapper.XPathMapper;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SheetUtil {
    private static Pattern dispimagPattern = Pattern.compile(".*DISPIMG\\(\"(ID_[\\dA-F]{32})\".*");
    /**
     * 是否日期单元格
     * @param cell
     * @return
     */
    public static boolean isDateCell(Cell cell){

        if (null != cell) {
            CellType cellType = cell.getCellTypeEnum();
            //支持公式单元格
            if (cellType == CellType.FORMULA) {
                cellType = cell.getCachedFormulaResultTypeEnum();
            }
            if(cellType == CellType.NUMERIC && HSSFDateUtil.isCellDateFormatted(cell)){
                return true;
            }
        }
        return false;
    }

    /**
     * 获取单元格的值，字符串
     * @param cell
     * @return
     */
    public static String getCellStringValue(Cell cell){
        String dateFormat = "yyyy-MM-dd HH:mm:ss";
        if (null != cell) {
            String str = null;
            CellType cellType = cell.getCellTypeEnum();
            //支持公式单元格
            if (cellType == CellType.FORMULA) {
                cellType = cell.getCachedFormulaResultTypeEnum();
            }
            switch (cellType) {
                case NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        str = StringUtil.format(cell.getDateCellValue(), dateFormat);
                    } else {
                        BigDecimal bd = new BigDecimal(String.valueOf(cell.getNumericCellValue()));
                        str = bd.stripTrailingZeros().toPlainString();
                    }
                    break;
                case BOOLEAN:
                    str = String.valueOf(cell.getBooleanCellValue());
                    break;
                case ERROR:
                    throw new RuntimeException("单元格为错误值");
                case STRING:
                default:
                    str = cell.getStringCellValue();
                    break;
            }
            return str;
        }
        return null;
    }

    public static byte[] getCellImageBytes(XSSFWorkbook workbook, Cell cell) {
        if (cell.getCellType() == CellType.FORMULA && cell.getCellFormula().contains("DISPIMG")) {
            Matcher matcher = dispimagPattern.matcher(cell.getCellFormula());
            if (!matcher.find()) {
                throw new RuntimeException("找不到ID");
            }
            String id = matcher.group(1);

            try {
                PackagePart cellimagesPart = workbook.getPackage().getParts().stream().filter(m -> m.getPartName().getName().equals("/xl/cellimages.xml")).findFirst().orElse(null);
                if (cellimagesPart == null) {
                    throw new RuntimeException("找不到图片");
                }
                XmlObject xmlObject = XmlObject.Factory.parse(cellimagesPart.getInputStream());
                CellImages cellImages = XPathMapper.parse(xmlObject.xmlText(), CellImages.class);
                PackagePart cellimagesRelsPart = workbook.getPackage().getParts().stream().filter(m -> m.getPartName().getName().equals("/xl/_rels/cellimages.xml.rels")).findFirst().orElse(null);
                if (cellimagesRelsPart == null) {
                    throw new RuntimeException("找不到图片");
                }
                XmlObject xmlObject2 = XmlObject.Factory.parse(cellimagesRelsPart.getInputStream());
                CellImagesRels cellImagesRels = XPathMapper.parse(xmlObject2.xmlText(), CellImagesRels.class);
                List<? extends PictureData> allPictures = workbook.getAllPictures();
                String rId = cellImages.getCellImageList().stream().filter(m -> m.getId().equals(id)).map(m -> m.getRId()).findFirst().orElse(null);
                if (rId == null) {
                    throw new RuntimeException("找不到图片");
                }
                String target = cellImagesRels.getCellImageRelsList().stream().filter(m -> m.getRId().equals(rId)).map(m -> m.getTarget()).findFirst().orElse(null);
                if (target == null) {
                    throw new RuntimeException("找不到图片");
                }
                byte[] bytes = allPictures.stream().filter(m -> ((XSSFPictureData) m).getPackagePart().getPartName().getName().equals("/xl/" + target)).map(m -> ((XSSFPictureData) m).getData()).findFirst().orElse(null);
                return bytes;

            } catch (XmlException e) {
                throw new RuntimeException(e);
            } catch (IOException e) {
                throw new RuntimeException(e);
            } catch (InvalidFormatException e) {
                throw new RuntimeException(e);
            }
        } else {
            throw new RuntimeException("非单元格图片");
        }
    }

}
