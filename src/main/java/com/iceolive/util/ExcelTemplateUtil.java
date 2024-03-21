package com.iceolive.util;

import jdk.nashorn.api.scripting.ScriptObjectMirror;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.script.Invocable;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class ExcelTemplateUtil {
    static Pattern tplReg = Pattern.compile("\\$\\{(.*?)}");
    static Pattern varReg = Pattern.compile("([$_a-zA-Z0-9\\.\\[\\]]+)$");

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
        ScriptEngineManager manager = new ScriptEngineManager();
        ScriptEngine js = manager.getEngineByName("JavaScript");
        if (variables != null) {
            for (String key : variables.keySet()) {
                js.put(key, variables.get(key));
            }
        }
        try {
            String func = cmd;
            if (func.trim().startsWith("return")) {
                func = "var _$result = function(){" + func + " }";
            } else {
                func = "var _$result = function(){return " + func + " }";
            }
            js.eval(func);
            Invocable inv = (Invocable) js;
            Object val = inv.invokeFunction("_$result");
            if (val instanceof ScriptObjectMirror) {
                val = toObject((ScriptObjectMirror) val);
            }
            if (val == null) {
                log.debug("JS: " + cmd + " => null");
            } else {
                log.debug("JS: " + cmd + " => " + val.toString());
            }
            return val;
        } catch (ScriptException | NoSuchMethodException e) {
            throw new RuntimeException(e);
        }
    }

    private static Object toObject(ScriptObjectMirror mirror) {
        if (mirror.isEmpty()) {
            return null;
        }
        if (mirror.isArray()) {
            List<Object> list = new ArrayList<>();
            for (Map.Entry<String, Object> entry : mirror.entrySet()) {
                Object result = entry.getValue();
                if (result instanceof ScriptObjectMirror) {
                    list.add(toObject((ScriptObjectMirror) result));
                } else {
                    list.add(result);
                }
            }
            return list;
        }

        Map<String, Object> map = new HashMap<>();
        for (Map.Entry<String, Object> entry : mirror.entrySet()) {
            Object result = entry.getValue();
            if (result instanceof ScriptObjectMirror) {
                map.put(entry.getKey(), toObject((ScriptObjectMirror) result));
            } else {
                map.put(entry.getKey(), result);
            }
        }
        return map;
    }

    public static void fillData(XSSFWorkbook workbook, Map<String, Object> variables ) {
        try {
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getPhysicalNumberOfRows();
            boolean hasCreateRow = false;
            for (int r = rowCount - 1; r >= 0; r--) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c <= row.getLastCellNum(); c++) {

                        Cell cell = row.getCell(c);
                        if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
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
                                    List jsonArray = (List) eval(listName, variables);
                                    if (jsonArray == null || jsonArray.size() == 0) {
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
                        if (cell != null && cell.getCellTypeEnum() == CellType.STRING) {
                            replaceParagraph(cell, variables);
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void replaceParagraph(Cell cell, Map<String, Object> variables) {
        if (cell == null || cell.getCellTypeEnum() != CellType.STRING) {
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
