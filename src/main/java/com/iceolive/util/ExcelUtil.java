package com.iceolive.util;

import com.iceolive.util.annotation.ExcelColumn;
import com.iceolive.util.model.ImportResult;
import com.iceolive.util.model.ValidateResult;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ReflectionUtils;

import javax.validation.ConstraintViolation;
import javax.validation.Valid;
import javax.validation.Validation;
import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;
import java.util.function.Function;

/**
 * @author wangmianzhe
 */
@SuppressWarnings("unchecked")
public class ExcelUtil {
    /**
     * 根据注解验证对象
     *
     * @param obj 验证的对象
     * @return 返回验证列表
     */
    private static List<ValidateResult> validate(@Valid Object obj) {
        List<ValidateResult> result = new ArrayList<>();
        Set<ConstraintViolation<@Valid Object>> validateSet = Validation.buildDefaultValidatorFactory()
                .getValidator()
                .validate(obj, new Class[0]);
        if (!CollectionUtils.isEmpty(validateSet)) {
            validateSet.stream().forEach((v) -> {
                String msg = v.getMessage();
                if (StringUtil.isEmpty(msg)) {
                    msg = "参数输入有误";
                }
                result.add(new ValidateResult(v.getPropertyPath().toString(), msg));
            });


        }
        return result;
    }


    /**
     * 导入excel
     *
     * @param filepath           excel文件路径
     * @param clazz              中间类类型
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类
     * @return 返回导入结果
     */
    public static <T> ImportResult importExcel(
            String filepath, Class<T> clazz,
            boolean faultTolerant,
            Function<T, List<ValidateResult>> customValidateFunc,
            Function<T, Boolean> importFunc) {
        FileInputStream stream;
        byte[] bytes = null;
        try {
            stream = new FileInputStream(filepath);
            int len = stream.available();
            bytes = new byte[len];
            stream.read(bytes);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return importExcel(bytes, clazz, faultTolerant, customValidateFunc, importFunc);
    }

    /**
     * 导入excel
     *
     * @param bytes              excel文件的字节数组
     * @param clazz              中间类类型
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类
     * @return 返回导入结果
     */
    public static <T> ImportResult<T> importExcel(
            byte[] bytes, Class<T> clazz,
            boolean faultTolerant,
            Function<T, List<ValidateResult>> customValidateFunc,
            Function<T, Boolean> importFunc) {
        ImportResult<T> result = new ImportResult<T>();
        result.setErrors(new ArrayList<>());
        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes));
        } catch (Exception e1) {
            try {
                workbook = new HSSFWorkbook(new ByteArrayInputStream(bytes));
            } catch (Exception e2) {
                throw new RuntimeException(e2);
            }
        }
        Sheet sheet = workbook.getSheetAt(0);
        //列序号和字段的map
        Map<Integer, Field> headMap = getHeadMap(sheet, clazz);
        //设置总记录数
        result.setTotalCount(sheet.getLastRowNum());
        Map<Integer, T> list = new LinkedHashMap<>();
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (null != row) {
                T obj = null;
                try {
                    obj = clazz.newInstance();
                } catch (IllegalAccessException | InstantiationException e) {
                    throw new RuntimeException(e);
                }
                boolean validate = true;
                for (Integer c : headMap.keySet()) {
                    Cell cell = row.getCell(c);
                    Field field = headMap.get(c);
                    if (null != cell) {
                        String str = null;
                        switch (cell.getCellTypeEnum()) {
                            case NUMERIC:
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    str = String.valueOf(cell.getDateCellValue());
                                } else {
                                    BigDecimal bd = new BigDecimal(String.valueOf(cell.getNumericCellValue()));
                                    str = bd.stripTrailingZeros().toPlainString();
                                }
                                break;
                            case BOOLEAN:
                                str = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case STRING:
                            default:
                                str = cell.getStringCellValue();
                                break;
                        }

                        try {
                            Object value = StringUtil.parse(str, field.getType());
                            ReflectionUtils.makeAccessible(field);
                            ReflectionUtils.setField(field, obj, value);
                        } catch (Exception e) {
                            validate = false;
                            ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                            errorMessage.setRow(row.getRowNum());
                            errorMessage.setCell(cell.getAddress().formatAsString());
                            errorMessage.setMessage("类型转换错误");
                            result.getErrors().add(errorMessage);
                        }
                    }
                }
                List<ValidateResult> validateResults = validate(obj);
                validate = isValidate(result, headMap, row, validate, validateResults);

                if (customValidateFunc != null) {
                    List<ValidateResult> customValidateResults = customValidateFunc.apply(obj);
                    validate = isValidate(result, headMap, row, validate, customValidateResults);
                }
                if (validate) {
                    list.put(row.getRowNum(), obj);
                }
            }

        }
        if (list.size() > 0) {
            if (faultTolerant || result.getErrors().size() == 0) {
                //如果容错模式或是验证全部通过
                if (importFunc != null) {
                    //如果有导入函数
                    for (Map.Entry<Integer, T> m : list.entrySet()) {
                        try {
                            if (Boolean.TRUE.equals(importFunc.apply(m.getValue()))) {
                                result.getSuccesses().put(m.getKey(), m.getValue());
                            } else {
                                ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                                errorMessage.setRow(m.getKey());
                                errorMessage.setMessage("未抛异常的错误");
                                result.getErrors().add(errorMessage);
                                //非容错模式，退出循环
                                if (!faultTolerant) {

                                    break;
                                }
                            }
                        } catch (Exception e) {
                            ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                            errorMessage.setRow(m.getKey());
                            errorMessage.setMessage(e.getMessage());
                            result.getErrors().add(errorMessage);
                            //非容错模式，退出循环
                            if (!faultTolerant) {
                                break;
                            }
                        }
                    }
                } else {
                    //没有导入函数
                    result.setSuccesses(list);
                }
            }
        }

        return result;

    }

    /**
     * 获取列序号和字段的对应关系
     *
     * @param sheet
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> Map<Integer, Field> getHeadMap(Sheet sheet, Class<T> clazz) {
        //列序号和字段的map
        Map<Integer, Field> headMap = new HashMap<>();
        //获取字段和列序号的对应关系
        if (sheet.getLastRowNum() > 0) {
            Row row = sheet.getRow(0);
            for (int c = 0; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (null != cell) {
                    String title = cell.getStringCellValue();
                    for (Field field : clazz.getDeclaredFields()) {
                        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                        if (excelColumn != null) {
                            if (excelColumn.value().equals(title)) {
                                headMap.put(c, field);
                            }
                        }
                    }

                }
            }
        }
        return headMap;
    }

    /**
     * 根据验证结果填充错误信息
     *
     * @param result
     * @param headMap
     * @param row
     * @param validate
     * @param validateResults
     * @return
     */
    private static boolean isValidate(ImportResult result, Map<Integer, Field> headMap, Row row, boolean validate, List<ValidateResult> validateResults) {
        if (validateResults != null && !validateResults.isEmpty()) {
            validate = false;
            for (ValidateResult v : validateResults) {
                for (Map.Entry<Integer, Field> m : headMap.entrySet()) {
                    if (m.getValue().getName().equals(v.getFieldName())) {
                        ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                        errorMessage.setRow(row.getRowNum());
                        errorMessage.setCell(new CellAddress(row.getRowNum(), m.getKey()).toString());
                        errorMessage.setMessage(v.getMessage());
                        result.getErrors().add(errorMessage);
                        break;
                    }
                }
            }

        }
        return validate;
    }


}
