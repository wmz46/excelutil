package com.iceolive.util;


import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.el.ExpressionFactory;
import javax.el.StandardELContext;
import javax.el.ValueExpression;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author wmz
 */
@Slf4j
public class WordTemplateUtil {

    static Pattern tplReg = Pattern.compile("\\$\\{(.*?)}");
    static Pattern varReg = Pattern.compile("([$_a-zA-Z0-9.\\[\\]]+)$");


    public static byte[] doc2bytes(XWPFDocument document) {
        if (document != null) {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            try {
                document.write(baos);
                baos.close();
            } catch (IOException e) {
                log.error("导出word异常", e);
            }
            return baos.toByteArray();
        } else {
            return null;
        }

    }

    public static void save(XWPFDocument document, String path) {
        byte[] bytes = doc2bytes(document);
        try (OutputStream out = new BufferedOutputStream(new FileOutputStream(path, false))) {
            out.write(bytes);
        } catch (IOException e) {
            log.error("写入文件异常", e);
            throw new RuntimeException("写入文件异常", e);
        }
    }

    public static void fillData(XWPFDocument document, Map<String, Object> variables) {
        try {
            format(document);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            //段落判断是否有[]，有则循环段落
            //这里必须使用for-i 循环，不允许用forIn代替，否则会报错
            for (int i = 0; i < paragraphs.size(); i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                Matcher matcher = tplReg.matcher(paragraph.getText());
                if (matcher.find()) {
                    String tpl = matcher.group(1);
                    if (tpl.contains("[]")) {
                        Matcher matcher1 = varReg.matcher(tpl.split("\\[]")[0]);
                        if (!matcher1.find()) {
                            continue;
                        }
                        String listName = matcher1.group(1);
                        List<?> jsonArray = (List<?>) eval(listName, variables);
                        if (jsonArray == null || jsonArray.isEmpty()) {
                            //如果列表为空，则清除整个段落
                            document.removeBodyElement(document.getPosOfParagraph(paragraph));
                        } else {
                            if (jsonArray.size() > 1) {
                                //如果列表大于1行，则根据列表长度-1添加行，添加行操作只操作一次
                                for (int num = jsonArray.size() - 1; num > 0; num--) {
                                    XWPFParagraph newPara = insertCloneParagraph(document, paragraph);

                                    replaceListParagraph(newPara, listName, num);
                                }
                            }

                            replaceListParagraph(paragraph, listName, 0);
                        }
                    }
                }
            }
            //表格判断是否有[],有则循环表格
            Iterator<XWPFTable> itTable = document.getTablesIterator();
            while (itTable.hasNext()) {
                XWPFTable table = itTable.next();
                //获得表格总行数
                int rowCount = table.getNumberOfRows();
                //遍历表格的每一行

                boolean hasCreateRow = false;
                for (int r = 0; r < rowCount; r++) {
                    XWPFTableRow row = table.getRow(r);
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (int c = 0; c < cells.size(); c++) {
                        XWPFTableCell cell = row.getCell(c);
                        Matcher matcher = tplReg.matcher(cell.getText());
                        if (matcher.find()) {
                            String tpl = matcher.group(1);
                            if (tpl.contains("[]")) {
                                Matcher matcher1 = varReg.matcher(tpl.split("\\[]")[0]);
                                if (!matcher1.find()) {
                                    continue;
                                }
                                String listName = matcher1.group(1);

                                List<?> jsonArray = (List<?>) eval(listName, variables);
                                if (jsonArray == null || jsonArray.isEmpty()) {
                                    //如果列表为空，则清除整个单元格
                                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                                        for (XWPFRun run : paragraph.getRuns()) {
                                            run.setText("", 0);
                                        }
                                    }
                                } else {
                                    if (jsonArray.size() > 1 && !hasCreateRow) {
                                        //如果列表大于1行，则根据列表长度-1添加行，添加行操作只操作一次
                                        for (int num = 1; num < jsonArray.size(); num++) {
                                            addCloneRow(table, row, r + 1);
                                        }
                                        hasCreateRow = true;
                                    }
                                    //填充列表下标
                                    for (int num = 0; num < jsonArray.size(); num++) {
                                        XWPFTableCell tmpCell = table.getRow(r + num).getCell(c);
                                        for (XWPFParagraph paragraph : tmpCell.getParagraphs()) {
                                            replaceListParagraph(paragraph, listName, num);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }


            //替换段落变量
            for (int i = document.getParagraphs().size() - 1; i >= 0; i--) {
                XWPFParagraph paragraph = document.getParagraphs().get(i);

                XmlCursor cursor = paragraph.getCTP().newCursor();
                cursor.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//*/w:txbxContent/w:p/w:r");
                List<XmlObject> ctrsintxtbx = new ArrayList<>();
                while (cursor.hasNextSelection()) {
                    cursor.toNextSelection();
                    XmlObject obj = cursor.getObject();
                    ctrsintxtbx.add(obj);
                }
                for (XmlObject obj : ctrsintxtbx) {
                    CTR ctr = CTR.Factory.parse(obj.xmlText());
                    //CTR ctr = CTR.Factory.parse(obj.newInputStream());
                    XWPFRun run = new XWPFRun(ctr, (IRunBody) paragraph);
                    replaceRun(run, variables);
                    obj.set(run.getCTR());
                }


                replaceParagraph(paragraph, variables);


            }

            // 替换表格中的指定文字
            itTable = document.getTablesIterator();
            while (itTable.hasNext()) {
                XWPFTable table = itTable.next();
                //获得表格总行数
                int count = table.getNumberOfRows();
                //遍历表格的每一行
                for (int i = 0; i < count; i++) {
                    //获得表格的行
                    XWPFTableRow row = table.getRow(i);
                    //在行元素中，获得表格的单元格
                    List<XWPFTableCell> cells = row.getTableCells();
                    //遍历单元格
                    for (XWPFTableCell cell : cells) {
                        List<XWPFParagraph> cellParagraphs = cell.getParagraphs();
                        for (XWPFParagraph cellParagraph : cellParagraphs) {
                            replaceParagraph(cellParagraph, variables);
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error(e.getMessage(),e);
        }

    }

    private static void replaceListParagraph(XWPFParagraph paragraph, String listName, int num) {
        List<XWPFRun> runs = paragraph.getRuns();
        int length = runs.size();
        for (int i = length - 1; i >= 0; i--) {
            String text = runs.get(i).getText(runs.get(i).getTextPosition());
            if (text != null) {
                String newText = text.replace(listName + "[]", listName + "[" + num + "]");
                runs.get(i).setText(newText, 0);
            }

        }
    }

    private static void replaceRun(XWPFRun run, Map<String, Object> variables) {
        String text = run.getText(run.getTextPosition());
        Matcher matcher = tplReg.matcher(text);
        String value = text;
        while (matcher.find()) {
            try {
                Object val = eval(matcher.group(1), variables);
                String s = "";
                if (val != null) {
                    s = String.valueOf(val);
                }
                value = value.replace("${" + matcher.group(1) + "}", s);
                run.setText(value, 0);
            } catch (Exception e) {
                log.error(e.toString());
            }
        }
    }


    private static void replaceParagraph(XWPFParagraph paragraph, Map<String, Object> variables) {
        List<XWPFRun> runs = paragraph.getRuns();
        int length = runs.size();
        //由于可能增加换行，从后面开始循环
        for (int i = length - 1; i >= 0; i--) {
            String text = runs.get(i).getText(runs.get(i).getTextPosition());

            if (text != null) {
                Matcher matcher = tplReg.matcher(text);
                String value = text;
                while (matcher.find()) {
                    try {
                        Object val = eval(matcher.group(1), variables);
                        String s = "";
                        if (val != null) {
                            s = String.valueOf(val);
                        }
                        value = value.replace("${" + matcher.group(1) + "}", s);

                        if (value.contains("\n")) {
                            RunProps runProps = getRunProps(paragraph.getRuns().get(i));
                            paragraph.removeRun(i);
                            String[] arr = value.split("\n");
                            for (int j = arr.length - 1; j >= 0; j--) {
                                XWPFRun run = paragraph.insertNewRun(i);
                                setRunProps(run, runProps);
                                if (j != 0) {
                                    run.addBreak(BreakClear.ALL);
                                    if(paragraph.getFirstLineIndent()!=0){
                                        run.addTab();
                                    }
                                }
                                run.setText(arr[j]);
                            }
                        } else {
                            runs.get(i).setText(value, 0);
                        }
                    } catch (Exception e) {
                        log.error(e.toString());
                    }
                }
            }
        }
    }

    private static void format(XWPFDocument document) {
        try {
            Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
            //遍历段落
            while (itPara.hasNext()) {
                XWPFParagraph paragraph = itPara.next();
                List<XWPFRun> run = paragraph.getRuns();
                for (int i = run.size() - 1; i > 0; i--) {
                    if (isSameFormat(run.get(i - 1), run.get(i))) {
                        run.get(i - 1).setText(run.get(i - 1).getText(run.get(i - 1).getTextPosition()) + run.get(i).getText(run.get(i).getTextPosition()), 0);
                        paragraph.removeRun(i);
                    }
                }
            }
            Iterator<XWPFTable> itTable = document.getTablesIterator();
            while (itTable.hasNext()) {
                XWPFTable table = itTable.next();
                //获得表格总行数
                int count = table.getNumberOfRows();
                //遍历表格的每一行
                for (int i = 0; i < count; i++) {
                    //获得表格的行
                    XWPFTableRow row = table.getRow(i);
                    //在行元素中，获得表格的单元格
                    List<XWPFTableCell> cells = row.getTableCells();
                    //遍历单元格
                    for (XWPFTableCell cell : cells) {
                        List<XWPFParagraph> cellParagraphs = cell.getParagraphs();
                        for (XWPFParagraph cellParagraph : cellParagraphs) {
                            List<XWPFRun> run = cellParagraph.getRuns();
                            for (int j = run.size() - 1; j > 0; j--) {
                                if (isSameFormat(run.get(j - 1), run.get(j))) {
                                    run.get(j - 1).setText(run.get(j - 1).getText(run.get(j - 1).getTextPosition()) + run.get(j).getText(run.get(j).getTextPosition()), 0);
                                    cellParagraph.removeRun(j);
                                }
                            }
                        }

                    }
                }
            }
        } catch (Exception e) {
            log.error(e.getMessage(),e);
        }
    }

    private static RunProps getRunProps(XWPFRun run) {
        RunProps props = new RunProps();
        props.setBold(run.isBold());
        props.setColor(run.getColor());
        props.setItalic(run.isItalic());
        props.setUnderline(run.getUnderline());
        props.setStrikeThrough(run.isStrikeThrough());
        props.setDoubleStrikeThrough(run.isDoubleStrikeThrough());
        props.setSmallCaps(run.isSmallCaps());
        props.setCapitalized(run.isCapitalized());
        props.setShadow(run.isShadowed());
        props.setImprinted(run.isImprinted());
        props.setEmbossed(run.isEmbossed());
        props.setKerning(run.getKerning());
        props.setFontFamily(run.getFontFamily());
        props.setFontSize(run.getFontSizeAsDouble());
        props.setVerticalAlignment(run.getVerticalAlignment());
        props.setTextPosition(run.getTextPosition());
        return props;
    }

    private static void setRunProps(XWPFRun run, RunProps props) {
        run.setBold(props.isBold());
        run.setColor(props.getColor());
        run.setItalic(props.isItalic());
        run.setUnderline(props.getUnderline());
        run.setStrikeThrough(props.isStrikeThrough());
        run.setDoubleStrikethrough(props.isDoubleStrikeThrough());
        run.setSmallCaps(props.isSmallCaps());
        run.setCapitalized(props.isCapitalized());
        run.setShadow(props.isShadow());
        run.setImprinted(props.isImprinted());
        run.setEmbossed(props.isEmbossed());
        if(props.getVerticalAlignment()!=null) {
            run.setVerticalAlignment(props.getVerticalAlignment().toString().toLowerCase());
        }
        run.setKerning(props.getKerning());
        run.setFontFamily(props.getFontFamily());
        if (props.getFontSize() != null) {
            run.setFontSize(props.getFontSize());
        }
        run.setTextPosition(props.getTextPosition());
    }

    @Data
    private static class RunProps {
        private boolean bold;
        private String color;
        private boolean italic;
        private UnderlinePatterns underline;
        private boolean strikeThrough;
        private boolean doubleStrikeThrough;
        private boolean smallCaps;
        private boolean capitalized;
        private boolean shadow;
        private boolean imprinted;
        private boolean embossed;
        private int kerning;
        private STVerticalAlignRun.Enum verticalAlignment;
        private String fontFamily;
        private Double fontSize;
        private int textPosition;

    }


    private static XWPFRun addCloneRun(XWPFParagraph paragraph, XWPFRun run) {
        XWPFRun r = paragraph.createRun();
        r.setBold(run.isBold());
        r.setColor(run.getColor());
        r.setText(run.text(), 0);
        r.setItalic(run.isItalic());
        r.setUnderline(run.getUnderline());
        r.setStrikeThrough(run.isStrikeThrough());
        r.setDoubleStrikethrough(run.isDoubleStrikeThrough());
        r.setSmallCaps(run.isSmallCaps());
        r.setCapitalized(run.isCapitalized());
        r.setShadow(run.isShadowed());
        r.setImprinted(run.isImprinted());
        r.setEmbossed(run.isEmbossed());
        r.setKerning(run.getKerning());
        r.setFontFamily(run.getFontFamily());
        if(run.getFontSizeAsDouble()!=null) {
            r.setFontSize(run.getFontSizeAsDouble());
        }
        r.setTextPosition(run.getTextPosition());
        if (run.getCTR().getRPr() != null) {
            r.getCTR().setRPr(run.getCTR().getRPr());
        }
        return r;
    }

    private static XWPFTableRow addCloneRow(XWPFTable table, XWPFTableRow row, int pos) {
        XWPFTableRow newRow = table.insertNewTableRow(pos);
        newRow.setHeight(row.getHeight());
        for (XWPFTableCell tableCell : row.getTableCells()) {
            XWPFTableCell newCell = newRow.addNewTableCell();
            //新行需要清除第一个段落
            newCell.removeParagraph(0);
            cloneCell(tableCell, newCell);
            for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
                XWPFParagraph p = newCell.addParagraph();
                cloneParagraph(p, paragraph);
            }
        }
        return newRow;
    }

    private static void cloneCell(XWPFTableCell cell, XWPFTableCell newCell) {
        CTTc ctTc = cell.getCTTc();
        CTTcPr tcPr = ctTc.getTcPr();
        if (tcPr != null) {
            CTTcPr newTcPr = newCell.getCTTc().getTcPr();
            if (newTcPr == null) {
                newTcPr = newCell.getCTTc().addNewTcPr();
            }
            if (tcPr.getTcW() != null) {
                newTcPr.addNewTcW().setW(tcPr.getTcW().getW());
            }
            if (tcPr.getVAlign() != null) {
                newTcPr.addNewVAlign().setVal(tcPr.getVAlign().getVal());
            }

            if (tcPr.getTcBorders() != null) {
                newTcPr.setTcBorders(tcPr.getTcBorders());
            }
            if (tcPr.getGridSpan() != null) {
                newTcPr.addNewGridSpan().setVal(tcPr.getGridSpan().getVal());
            }
        }
    }

    private static void cloneParagraph(XWPFParagraph p, XWPFParagraph paragraph) {
        p.setAlignment(paragraph.getAlignment());
        p.setFontAlignment(paragraph.getFontAlignment());
        p.setVerticalAlignment(paragraph.getVerticalAlignment());
        p.setBorderTop(paragraph.getBorderTop());
        p.setBorderBottom(paragraph.getBorderBottom());
        p.setBorderLeft(paragraph.getBorderLeft());
        p.setBorderRight(paragraph.getBorderRight());
        p.setBorderBetween(paragraph.getBorderBetween());
        p.setPageBreak(paragraph.isPageBreak());
        p.setSpacingAfter(paragraph.getSpacingAfter());
        p.setSpacingAfterLines(paragraph.getSpacingAfterLines());
        p.setSpacingBefore(paragraph.getSpacingBefore());
        p.setSpacingBeforeLines(paragraph.getSpacingBeforeLines());
        p.setSpacingLineRule(paragraph.getSpacingLineRule());
        p.setSpacingBetween(paragraph.getSpacingBetween());
        p.setIndentationLeft(paragraph.getIndentationLeft());
        p.setIndentationRight(paragraph.getIndentationRight());
        p.setIndentationHanging(paragraph.getIndentationLeft());
        p.setIndentationFirstLine(paragraph.getIndentationFirstLine());
        p.setIndentFromLeft(paragraph.getIndentFromLeft());
        p.setIndentFromRight(paragraph.getIndentFromRight());
        p.setFirstLineIndent(paragraph.getFirstLineIndent());
        p.setStyle(paragraph.getStyle());
        if (paragraph.getCTP().getPPr() != null) {
            CTPPr newppr = p.getCTP().addNewPPr();
            if (paragraph.getCTP().getPPr().getJc() != null) {
                newppr.addNewJc().setVal(paragraph.getCTP().getPPr().getJc().getVal());
            }
            CTSpacing spacing = paragraph.getCTP().getPPr().getSpacing();
            //段落间距
            if (spacing != null) {
                CTSpacing newSpacing = newppr.addNewSpacing();
                newSpacing.setAfter(spacing.getAfter());
                newSpacing.setAfterAutospacing(spacing.getAfterAutospacing());
                newSpacing.setAfterLines(spacing.getAfterLines());
                newSpacing.setBefore(spacing.getBefore());
                newSpacing.setBeforeAutospacing(spacing.getBeforeAutospacing());
                newSpacing.setBeforeLines(spacing.getBeforeLines());
                newSpacing.setLine(newSpacing.getLine());
                newSpacing.setLineRule(spacing.getLineRule());
            }
            //段落缩进
            CTInd ind = paragraph.getCTP().getPPr().getInd();
            if (ind != null) {
                CTInd newInd = p.getCTP().getPPr().addNewInd();
                newInd.setFirstLine(ind.getFirstLine());
                newInd.setFirstLineChars(ind.getFirstLineChars());
                newInd.setHanging(ind.getHanging());
                newInd.setHangingChars(ind.getHangingChars());
                newInd.setLeft(ind.getLeft());
                newInd.setLeftChars(ind.getLeftChars());
                newInd.setRight(ind.getRight());
                newInd.setRightChars(ind.getRightChars());
            }

        }
        for (XWPFRun run : paragraph.getRuns()) {
            addCloneRun(p, run);
        }
    }

    private static XWPFParagraph insertCloneParagraph(XWPFDocument document, XWPFParagraph paragraph) {
        XmlCursor cursor = paragraph.getCTP().newCursor();
        //光标移到下一个段落
        cursor.toNextSibling();
        XWPFParagraph p = document.insertNewParagraph(cursor);
        cloneParagraph(p, paragraph);
        return p;

    }

    private static boolean isSameFormat(XWPFRun run1, XWPFRun run2) {
        if (ObjectUtils.notEqual(run1.isBold(), run2.isBold())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.isCapitalized(), run2.isCapitalized())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getCharacterSpacing(), run2.getCharacterSpacing())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getColor(), run2.getColor())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.isDoubleStrikeThrough(), run2.isDoubleStrikeThrough())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.isEmbossed(), run2.isEmbossed())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getFontFamily(), run2.getFontFamily())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getFontSizeAsDouble(), run2.getFontSizeAsDouble())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.isImprinted(), run2.isImprinted())) {
            return false;
        }

        if (ObjectUtils.notEqual(run1.isItalic(), run2.isItalic())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getKerning(), run2.getKerning())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.isShadowed(), run2.isShadowed())) {
            return false;
        }

        if (ObjectUtils.notEqual(run1.isSmallCaps(), run2.isSmallCaps())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.isStrikeThrough(), run2.isStrikeThrough())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getTextPosition(), run2.getTextPosition())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getUnderline(), run2.getUnderline())) {
            return false;
        }
        if (ObjectUtils.notEqual(run1.getCTR().getRPr().toString(), run2.getCTR().getRPr().toString())) {
            return false;
        }
        return run1.getText(0) != null && run2.getText(0) != null;
    }

    public static void saveXml(XWPFDocument document, String xmlPath) {
        String xmlString = document.getDocument().toString();

        String xml = xmlString.replaceAll("<xml-fragment.*?>", "").replaceAll("</xml-fragment>", "");
        xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                + "<?mso-application progid=\"Word.Document\"?>\n"
                + "<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:sl=\"http://schemas.microsoft.com/schemaLibrary/2003/core\"  xmlns:aml=\"http://schemas.microsoft.com/aml/2001/core\" xmlns:wx=\"http://schemas.microsoft.com/office/word/2003/auxHint\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:dt=\"uuid:C2F41010-65B3-11d1-A29F-00AA00C14882\" w:macrosPresent=\"no\" w:embeddedObjPresent=\"no\" w:ocxPresent=\"no\" xml:space=\"preserve\" xmlns:wpsCustomData=\"http://www.wps.cn/officeDocument/2013/wpsCustomData\">"
                + xml
                + "</w:wordDocument>";

        byte[] bytes = xml.getBytes(StandardCharsets.UTF_8);
        try (OutputStream out = new BufferedOutputStream(new FileOutputStream(xmlPath, false))) {
            out.write(bytes);
        } catch (IOException e) {
            log.error("写入文件异常", e);
            throw new RuntimeException("写入文件异常", e);
        }

    }

    public static void saveWordXml(XWPFDocument document, String xmlPath) {
        String xmlString = document.getDocument().toString();

        String xml = xmlString.replace("<xml-fragment", "<w:document").replace("</xml-fragment>", "</w:document>");
        xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<?mso-application progid=\"Word.Document\"?>\n" +
                "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" +
                "<pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\">\n" +
                "<pkg:xmlData>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n" +
                "</Relationships>\n" +
                "</pkg:xmlData>\n" +
                "</pkg:part>\n" +
                "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\">\n" +
                "<pkg:xmlData>" + xml + "</pkg:xmlData>\n" +
                "</pkg:part>\n" +
                "</pkg:package>";
        byte[] bytes = xml.getBytes(StandardCharsets.UTF_8);
        try (OutputStream out = new BufferedOutputStream(new FileOutputStream(xmlPath, false))) {
            out.write(bytes);
        } catch (IOException e) {
            log.error("写入文件异常", e);
            throw new RuntimeException("写入文件异常", e);
        }

    }

    /**
     * 加载xml或doc文件，xml最好是由该工具导出，不然会有一些格式丢失。
     *
     * @param filePath word文件路径
     * @return 返回word文档
     */
    public static XWPFDocument load(String filePath) {
        try {
            if (filePath.toLowerCase().endsWith(".docx")) {
                return new XWPFDocument(Files.newInputStream(Paths.get(filePath)));
            } else if (filePath.toLowerCase().endsWith(".xml")) {
                SAXReader saxReader = new SAXReader();
                Document document = saxReader.read(new File(filePath));
                Element element = document.getRootElement().elements().stream().filter(m -> "/word/document.xml".equals(m.attribute("name").getText())).findFirst().orElse(null);
                if (element != null) {
                    element = element.element("xmlData").element("document");
                } else {
                    element = document.getRootElement().elements().stream().findFirst().orElse(null);
                }
                assert element != null;
                element.setName("xml-fragment");
                XWPFDocument xwpfDocument = new XWPFDocument();
                CTDocument1 ctDocument1 = CTDocument1.Factory.parse(element.asXML());
                xwpfDocument.getDocument().set(ctDocument1);
                return xwpfDocument;

            } else {
                throw new RuntimeException("不支持的格式");
            }
        } catch (IOException | XmlException | DocumentException e) {
            throw new RuntimeException(e);
        }
    }

    public static Object eval(String cmd, Map<String, Object> variables) {
        ExpressionFactory factory = ExpressionFactory.newInstance();
        StandardELContext context = new StandardELContext(factory);
        if (variables != null) {
            for (Map.Entry<String, Object> entry : variables.entrySet()) {
                context.getVariableMapper().setVariable(entry.getKey(), factory.createValueExpression(variables.get(entry.getKey()), Object.class));
            }
        }
        ValueExpression expression = factory.createValueExpression(context, "${" + cmd + "}", Object.class);
        return expression.getValue(context);
    }


}
