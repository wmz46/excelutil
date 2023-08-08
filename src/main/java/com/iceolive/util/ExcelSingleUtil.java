package com.iceolive.util;

import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.exception.ImageOutOfBoundsException;
import com.iceolive.util.model.FieldInfo;
import com.iceolive.util.model.ImportResult;
import com.iceolive.util.model.ImportSingleResult;
import com.iceolive.util.model.ValidateResult;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

public class ExcelSingleUtil {
    public static ImportSingleResult importExcel(String filepath, List<FieldInfo> fieldInfos) {
        FileInputStream inputStream;
        try {
            inputStream = new FileInputStream(filepath);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return importExcel(inputStream, fieldInfos);
    }

    public static ImportSingleResult importExcel(InputStream inputStream, List<FieldInfo> fieldInfos) {
        ImportSingleResult result = new ImportSingleResult();
        result.setErrors(new ArrayList<>());
        Map<String, Object> obj = new HashMap<>();
        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(inputStream);

        } catch (Exception e1) {
            try {
                workbook = new HSSFWorkbook(inputStream);
            } catch (Exception e2) {
                throw new RuntimeException(e2);
            }
        }
        Sheet sheet = workbook.getSheetAt(0);
        String dateFormat = "yyyy-MM-dd HH:mm:ss";

        boolean validate = true;
        for (FieldInfo fieldInfo : fieldInfos) {
            String cellString = fieldInfo.getCellString();
            CellAddress cellAddress = new CellAddress(cellString);
            Row row = sheet.getRow(cellAddress.getRow());
            int c = cellAddress.getColumn();
            try {
                if (row != null) {
                    Cell cell = row.getCell(cellAddress.getColumn());
                    boolean isDateCell = SheetUtil.isDateCell(cell);
                    if (cell != null) {
                        String str = SheetUtil.getCellStringValue(cell);
                        Object value = null;
                        if (isDateCell || fieldInfo.getType() == ColumnType.DATETIME.getValue() || fieldInfo.getType() == ColumnType.DATE.getValue()) {
                            //特殊处理日期格式
                            if (!StringUtil.isBlank(str)) {
                                value = StringUtil.parse(str, dateFormat, Date.class);
                            }
                        } else if (fieldInfo.getType() == ColumnType.IMAGE.getValue()) {
                            value = SheetUtil.getCellImageBytes((XSSFWorkbook) workbook, cell);
                        } else if (fieldInfo.getType() == ColumnType.LONG.getValue()) {
                            value = StringUtil.parse(str, Long.class);
                        } else if (fieldInfo.getType() == ColumnType.DOUBLE.getValue()) {
                            value = StringUtil.parse(str, Double.class);
                        } else {
                            value = str;
                        }
                        obj.put(fieldInfo.getName(), value);
                    }
                }
            } catch (Exception e) {
                validate = false;
                ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                errorMessage.setRow(row.getRowNum());
                errorMessage.setCol(CellReference.convertNumToColString(c));
                errorMessage.setCell(new CellAddress(row.getRowNum(), c).formatAsString());
                if (e instanceof ImageOutOfBoundsException) {
                    errorMessage.setMessage(e.getMessage());
                } else {
                    errorMessage.setMessage("类型转换错误");

                }
                result.getErrors().add(errorMessage);
            }
        }
        List<ValidateResult> validateResults = ValidateUtil.validate(obj, fieldInfos);
        if(!CollectionUtils.isEmpty(validateResults)){
            validate = false;
            for (ValidateResult v : validateResults) {
                FieldInfo fieldInfo = fieldInfos.stream().filter(m -> m.getName().equals(v.getFieldName())).findFirst().orElse(null);
                if(fieldInfo!=null) {
                    CellAddress cellAddress = new CellAddress(fieldInfo.getCellString());
                    ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                    errorMessage.setRow(cellAddress.getRow());
                    errorMessage.setCol(CellReference.convertNumToColString(cellAddress.getColumn()));
                    errorMessage.setCell(cellAddress.toString());
                    errorMessage.setMessage(v.getMessage());
                    result.getErrors().add(errorMessage);
                }
            }
        }

        if(validate) {
            result.setSuccess(obj);
        }
        return result;
    }
}
