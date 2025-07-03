package com.iceolive.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.el.ExpressionFactory;
import javax.el.StandardELContext;
import javax.el.ValueExpression;
import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author wmz
 */
@Slf4j
public class ExcelTemplateUtil {
    static Pattern tplReg = Pattern.compile("\\$\\{(.*?)}");
    static Pattern varReg = Pattern.compile("([$_a-zA-Z0-9.\\[\\]]+)$");

    public static XSSFWorkbook load(String filePath) {
        try {
            return new XSSFWorkbook(filePath);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static byte[] xlsx2bytes(XSSFWorkbook document) {
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

    public static void save(XSSFWorkbook workbook, String path) {
        byte[] bytes = xlsx2bytes(workbook);
        try (OutputStream out = new BufferedOutputStream(new FileOutputStream(path, false))) {
            out.write(bytes);
        } catch (IOException e) {
            log.error("写入文件异常", e);
            throw new RuntimeException("写入文件异常", e);
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

    public static void fillData(XSSFWorkbook workbook, Map<String, Object> variables ) {
          fillData(workbook,0,variables);
    }
    public static void fillData(XSSFWorkbook workbook,int sheetIndex, Map<String, Object> variables ) {
        try {
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            int rowCount = sheet.getPhysicalNumberOfRows();
            boolean hasCreateRow = false;
            for (int r = rowCount - 1; r >= 0; r--) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c <= row.getLastCellNum(); c++) {

                        Cell cell = row.getCell(c);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String text = cell.getStringCellValue();
                            Matcher matcher = tplReg.matcher(text);
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
                                        cell.setCellValue("");
                                    } else {
                                        if (jsonArray.size() > 1 && !hasCreateRow) {
                                            //如果列表大于1行，则根据列表长度-1添加行，添加行操作只操作一次

                                            for (int num = jsonArray.size() - 1; num > 0; num--) {
                                                if (r + 1 <= sheet.getLastRowNum()) {
                                                    sheet.shiftRows(r + 1, sheet.getLastRowNum(), 1);
                                                }
                                                sheet.createRow(r + 1);
                                                sheet.copyRows(r, r, r + 1, new CellCopyPolicy());

                                            }
                                            hasCreateRow = true;

                                        }
                                        //填充列表下标
                                        for (int num = 0; num < jsonArray.size(); num++) {
                                            Cell tmpCell = sheet.getRow(r + num).getCell(c);
                                            replaceListParagraph(tmpCell, listName, num);
                                        }

                                    }
                                }
                            }
                        }
                    }
                }
            }
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c <= row.getLastCellNum(); c++) {
                        Cell cell = row.getCell(c);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            replaceParagraph(cell, variables);
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error(e.getMessage(),e);
        }
    }

    private static void replaceParagraph(Cell cell, Map<String, Object> variables) {
        if (cell == null || cell.getCellType() != CellType.STRING) {
            return;
        }
        String text = cell.getStringCellValue();
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
                    cell.setCellValue(value);
                } catch (Exception e) {
                    log.error(e.toString());
                }
            }
        }
    }

    private static void replaceListParagraph(Cell cell, String listName, int num) {
        String text = cell.getStringCellValue();
        String newText = text.replace(listName + "[]", listName + "[" + num + "]");
        cell.setCellValue(newText);

    }

}
