package com.iceolive.util;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.github.pjfanning.xlsx.StreamingReader;
import com.iceolive.util.annotation.ExcelColumn;
import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.exception.ImageOutOfBoundsException;
import com.iceolive.util.model.*;
import com.networknt.schema.JsonSchema;
import com.networknt.schema.JsonSchemaFactory;
import com.networknt.schema.SpecVersionDetector;
import com.networknt.schema.ValidationMessage;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import javax.validation.ConstraintViolation;
import javax.validation.Valid;
import javax.validation.Validation;
import javax.validation.Validator;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Time;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;
import java.util.function.Function;

/**
 * @author wangmianzhe
 */
@Slf4j
@SuppressWarnings("unchecked")
public class ExcelUtil {

    private static Validator validator = null;

    private static Validator getValidatorInstance() {
        if (validator == null) {
            createValidatorInstance();
        }
        return validator;
    }

    private static synchronized Validator createValidatorInstance() {
        if (validator == null) {
            validator = Validation.buildDefaultValidatorFactory()
                    .getValidator();
        }
        return validator;
    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            String filepath, Class<T> clazz,
            boolean faultTolerant) {
        return importExcel(filepath, clazz, faultTolerant, 1, null, null);

    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            String filepath, Class<T> clazz,
            boolean faultTolerant, int startRow) {
        return importExcel(filepath, clazz, faultTolerant, startRow, null, null);

    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            String filepath, Class<T> clazz,
            boolean faultTolerant,
            Function<T, Boolean> importFunc) {
        return importExcel(filepath, clazz, faultTolerant, 1, null, importFunc);

    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            String filepath, Class<T> clazz,
            boolean faultTolerant, int startRow,
            Function<T, Boolean> importFunc) {
        return importExcel(filepath, clazz, faultTolerant, startRow, null, importFunc);
    }

    /**
     * 导入excel
     *
     * @param filepath           excel文件路径
     * @param clazz              中间类类型
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            String filepath, Class<T> clazz,
            boolean faultTolerant,
            Function<T, List<ValidateResult>> customValidateFunc,
            Function<T, Boolean> importFunc) {
        return importExcel(filepath, clazz, faultTolerant, 1, customValidateFunc, importFunc);
    }


    /**
     * 导入excel
     *
     * @param filepath           excel文件路径
     * @param clazz              中间类类型
     * @param startRow           开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T>                中间类
     * @return 返回导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            String filepath, Class<T> clazz,
            boolean faultTolerant,
            int startRow,
            Function<T, List<ValidateResult>> customValidateFunc,
            Function<T, Boolean> importFunc) {
        FileInputStream inputStream;
        try {
            inputStream = new FileInputStream(filepath);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return importExcel(inputStream, clazz, faultTolerant, startRow, customValidateFunc, importFunc);
    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            InputStream inputStream, Class<T> clazz,
            boolean faultTolerant) {
        return importExcel(inputStream, clazz, faultTolerant, 0, null, null);
    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            InputStream inputStream, Class<T> clazz,
            boolean faultTolerant, int startRow) {
        return importExcel(inputStream, clazz, faultTolerant, startRow, null, null);
    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            InputStream inputStream, Class<T> clazz,
            boolean faultTolerant,
            Function<T, Boolean> importFunc) {
        return importExcel(inputStream, clazz, faultTolerant, 0, null, importFunc);
    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param clazz         中间类类型
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            InputStream inputStream, Class<T> clazz,
            boolean faultTolerant, int startRow,
            Function<T, Boolean> importFunc) {
        return importExcel(inputStream, clazz, faultTolerant, startRow, null, importFunc);
    }

    /**
     * 导入excel
     *
     * @param inputStream        excel文件的字节数组
     * @param clazz              中间类类型
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T> 中间类类型
     * @return 导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            InputStream inputStream, Class<T> clazz,
            boolean faultTolerant,
            Function<T, List<ValidateResult>> customValidateFunc,
            Function<T, Boolean> importFunc) {
        return importExcel(inputStream, clazz, faultTolerant, 0, customValidateFunc, importFunc);
    }

    /**
     * 导入excel
     *
     * @param inputStream        excel文件的字节数组
     * @param clazz              中间类类型
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow           开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @param <T>                中间类
     * @return 返回导入结果
     */
    @Deprecated
    public static <T> ImportResult<T> importExcel(
            InputStream inputStream, Class<T> clazz,
            boolean faultTolerant,
            int startRow,
            Function<T, List<ValidateResult>> customValidateFunc,
            Function<T, Boolean> importFunc) {
        return importExcel(ExcelImportConfig.<T>builder()
                .inputStream(inputStream)
                .clazz(clazz)
                .faultTolerant(faultTolerant)
                .startRow(startRow)
                .customValidateFunc(customValidateFunc)
                .importFunc(importFunc)
                .build());

    }

    public static <T> ImportResult<T> importExcel(ExcelImportConfig<T> config) {

        Class<T> clazz = config.getClazz();
        InputStream inputStream = config.getInputStream();
        if (inputStream == null && StringUtil.isNotBlank(config.getFilepath())) {
            try {
                inputStream = Files.newInputStream(Paths.get(config.getFilepath()));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        assert inputStream != null;

        int startRow = config.getStartRow();
        Function<T, List<ValidateResult>> customValidateFunc = config.getCustomValidateFunc();
        boolean faultTolerant = config.isFaultTolerant();
        Function<T, Boolean> importFunc = config.getImportFunc();
        ImportResult<T> result = new ImportResult<>();
        result.setErrors(new ArrayList<>());
        Workbook workbook;
        try {
            if (hasCellImageField(clazz)) {
                //如果有图片字段，则不使用StreamingWorkbook
                workbook = new XSSFWorkbook(inputStream);
            } else {
                workbook = StreamingReader.builder()
                        //缓存到内存中的行数，默认是10
                        .rowCacheSize(100)
                        //读取资源时，缓存到内存的字节大小，默认是1024
                        .bufferSize(4096)
                        //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
                        .open(inputStream);
            }
        } catch (Exception e1) {
            try {
                workbook = new HSSFWorkbook(inputStream);
            } catch (Exception e2) {
                throw new RuntimeException(e2);
            }
        }
        Sheet sheet = workbook.getSheetAt(config.getSheetIndex());
        //列序号和字段的map
        Map<Integer, List<Field>> headMap = new HashMap<>();
        if (config.isOnlyData()) {
            for (Field field : clazz.getDeclaredFields()) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                if (excelColumn != null) {
                    if (StringUtil.isNotBlank(excelColumn.colString())) {
                        int i = CellReference.convertColStringToIndex(excelColumn.colString());
                        if (headMap.containsKey(i)) {
                            headMap.get(i).add(field);
                        } else {
                            headMap.put(i, new ArrayList<Field>() {{
                                add(field);
                            }});
                        }
                    } else {
                        throw new RuntimeException("当onlyData=true时，@ExcelColumn必须设置colString");
                    }
                }
            }
        }
        Map<Integer, T> list = new LinkedHashMap<>();
        int totalCount = 0;
        int titleRowNum = startRow - 1;
        for (Row row : sheet) {
            if (row.getRowNum() < titleRowNum) {
                //小于标题行的抛弃
            } else if (row.getRowNum() == titleRowNum && !config.isOnlyData()) {
                headMap = getHeadMap(clazz, row);

            } else {
                totalCount++;
                boolean isEmptyRow = true;
                T obj;
                try {
                    obj = clazz.newInstance();
                } catch (IllegalAccessException | InstantiationException e) {
                    throw new RuntimeException(e);
                }
                boolean validate = true;
                for (Integer c : headMap.keySet()) {
                    Cell cell = row.getCell(c);
                    List<Field> fields = headMap.get(c);
                    //是否日期单元格
                    boolean isDateCell = SheetUtil.isDateCell(cell);
                    String dateFormat = "yyyy-MM-dd HH:mm:ss";
                    try {
                        if (null != cell) {
                            String str = SheetUtil.getCellStringValue(cell);
                            if (!StringUtil.isBlank(str)) {
                                isEmptyRow = false;
                            }
                            for (Field field : fields) {
                                Object value = null;
                                if (isDateCell && (field.getType().isAssignableFrom(Date.class) || field.getType().isAssignableFrom(LocalDateTime.class) || field.getType().isAssignableFrom(LocalDate.class))) {
                                    //特殊处理日期格式
                                    if (!StringUtil.isBlank(str)) {
                                        value = StringUtil.parse(str, dateFormat, field.getType());
                                    }
                                } else if (isDateCell && (field.getType().isAssignableFrom(LocalTime.class) || field.getType().isAssignableFrom(Time.class))) {
                                    //特殊处理日期格式
                                    if (!StringUtil.isBlank(str)) {
                                        value = StringUtil.parse(str, dateFormat, field.getType());
                                    }
                                } else if (field.getType().isAssignableFrom(boolean.class) || field.getType().isAssignableFrom(Boolean.class)) {
                                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                                    assert excelColumn != null;
                                    value = StringUtil.parseBoolean(str, excelColumn.trueString(), excelColumn.falseString(), field.getType());
                                } else if (field.getType().isArray() && field.getType().getComponentType().equals(byte.class)) {
                                    value = SheetUtil.getCellImageBytes((XSSFWorkbook) workbook, cell);
                                } else if (field.getType().isAssignableFrom(BufferedImage.class)) {
                                    value = ImageUtil.Bytes2Image(SheetUtil.getCellImageBytes((XSSFWorkbook) workbook, cell));
                                } else {
                                    value = StringUtil.parse(str, field.getType());
                                }

                                field.setAccessible(true);
                                field.set(obj, value);
                            }

                        } else {
                            //单元格为null，处理图片
                            for (Field field : fields) {
                                Object value = null;
                                if (field.getType().isArray() && field.getType().getComponentType().equals(byte.class)) {
                                    List<byte[]> floatImages = getFloatImagesBytes(sheet, row.getRowNum(), c);
                                    if (!CollectionUtils.isEmpty(floatImages)) {
                                        value = floatImages.get(0);
                                    }
                                } else if (field.getType().isAssignableFrom(BufferedImage.class)) {
                                    List<byte[]> floatImages = getFloatImagesBytes(sheet, row.getRowNum(), c);
                                    if (!CollectionUtils.isEmpty(floatImages)) {
                                        value = ImageUtil.Bytes2Image(floatImages.get(0));

                                    }
                                } else if (field.getType().isAssignableFrom(List.class)) {
                                    ParameterizedType genericType = (ParameterizedType) field.getGenericType();
                                    if (genericType.getActualTypeArguments()[0] == BufferedImage.class) {
                                        List<byte[]> floatImages = getFloatImagesBytes(sheet, row.getRowNum(), c);
                                        value = new ArrayList<>();
                                        for (byte[] floatImage : floatImages) {
                                            ((List<Object>) value).add(ImageUtil.Bytes2Image(floatImage));
                                        }
                                    } else if (genericType.getActualTypeArguments()[0] == byte[].class) {
                                        value = getFloatImagesBytes(sheet, row.getRowNum(), c);
                                    }
                                }
                                field.setAccessible(true);
                                field.set(obj, value);
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
                if (isEmptyRow) {
                    //空行时总数-1，且不校验
                    totalCount--;
                    continue;
                }

                List<ValidateResult> validateResults = validate(obj);
                validate = isValidate(result, headMap, row, validate, validateResults, clazz);

                if (customValidateFunc != null) {
                    List<ValidateResult> customValidateResults = customValidateFunc.apply(obj);
                    validate = isValidate(result, headMap, row, validate, customValidateResults, clazz);
                }
                if (validate) {
                    list.put(row.getRowNum(), obj);
                }
            }
        }
        //设置总记录数
        result.setTotalCount(totalCount);
        if (!list.isEmpty()) {
            if (faultTolerant || result.getErrors().isEmpty()) {
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
     * @param clazz 中间类类型
     * @param row   标题行
     * @param <T> 中间类类型
     * @return 导入列序号和字段的对应关系
     */
    private static <T> Map<Integer, List<Field>> getHeadMap(Class<T> clazz, Row row) {
        //列序号和字段的map
        Map<Integer, List<Field>> headMap = new HashMap<>();
        for (int c = 0; c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (null != cell) {
                String title = cell.getStringCellValue();
                for (Field field : clazz.getDeclaredFields()) {
                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn != null) {
                        if (StringUtil.isNotEmpty(excelColumn.value())) {
                            if (excelColumn.value().equals(title)) {
                                if (headMap.containsKey(c)) {
                                    headMap.get(c).add(field);
                                } else {
                                    List<Field> fieldList = new ArrayList<>();
                                    fieldList.add(field);
                                    headMap.put(c, fieldList);
                                }
                            }
                        } else {
                            //如果ExcelColumn不指定名称，则使用字段名匹配
                            if (field.getName().equals(title)) {
                                if (headMap.containsKey(c)) {
                                    headMap.get(c).add(field);
                                } else {
                                    List<Field> fieldList = new ArrayList<>();
                                    fieldList.add(field);
                                    headMap.put(c, fieldList);
                                }
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
     * @param result 导入结果
     * @param headMap 行map
     * @param row 当前行
     * @param validate 验证是否通过
     * @param validateResults 验证信息
     * @return 验证是否通过
     */
    private static boolean isValidate(ImportResult<?> result, Map<Integer, List<Field>> headMap, Row row, boolean validate, List<ValidateResult> validateResults, Class<?> clazz) {
        if (validateResults != null && !validateResults.isEmpty()) {
            validate = false;
            for (ValidateResult v : validateResults) {
                //错误是否在单元格内
                boolean errorInCell = false;
                for (Map.Entry<Integer, List<Field>> m : headMap.entrySet()) {
                    List<Field> fields = m.getValue();
                    boolean stop = false;
                    for (Field field : fields) {
                        if (field.getName().equals(v.getFieldName())) {
                            ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                            errorMessage.setRow(row.getRowNum());
                            errorMessage.setCol(CellReference.convertNumToColString(m.getKey()));
                            errorMessage.setCell(new CellAddress(row.getRowNum(), m.getKey()).toString());
                            errorMessage.setMessage(v.getMessage());
                            result.getErrors().add(errorMessage);
                            stop = true;
                            errorInCell = true;
                            break;
                        }
                    }
                    if (stop) {
                        break;
                    }
                }
                if (!errorInCell) {
                    String fieldName = v.getFieldName();
                    String columnName = fieldName;
                    Field field = Arrays.stream(clazz.getDeclaredFields()).filter(m -> m.getName().equals(fieldName)).findFirst().orElse(null);
                    if (field != null) {
                        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                        if (excelColumn != null) {
                            String title = excelColumn.value();
                            if (StringUtil.isNotEmpty(title)) {
                                columnName = title;
                            }
                        }
                    }
                    //如果错误不在单元格内，不
                    ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                    errorMessage.setRow(row.getRowNum());
                    errorMessage.setMessage(v.getMessage() + "\n请检查[" + columnName + "]列是否存在");
                    result.getErrors().add(errorMessage);
                }
            }

        }
        return validate;
    }


    /**
     * 根据注解验证对象
     *
     * @param obj 验证的对象
     * @return 返回验证列表
     */
    public static List<ValidateResult> validate(@Valid Object obj) {
        List<ValidateResult> result = new ArrayList<>();
        Set<ConstraintViolation<@Valid Object>> validateSet = getValidatorInstance()
                .validate(obj);
        if (!CollectionUtils.isEmpty(validateSet)) {
            validateSet.forEach((v) -> {
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
     * json-schema验证
     *
     * @param schemaJson json结构
     * @param obj 对象
     * @return 验证结果信息
     */
    public static List<ValidateResult> jsonSchemaValidate(String schemaJson, Object obj) {
        List<ValidateResult> result = new ArrayList<>();
        try {
            ObjectMapper mapper = new ObjectMapper();
            JsonNode schemaNode = mapper.readTree(schemaJson);
            JsonSchema schema = JsonSchemaFactory.getInstance(SpecVersionDetector.detect(schemaNode)).getSchema(schemaNode);
            JsonNode node = mapper.readTree(mapper.writeValueAsString(obj));
            Set<ValidationMessage> errors = schema.validate(node);
            for (ValidationMessage error : errors) {
                String msg = error.getMessage();
                String fieldName = error.getPath();
                int split = msg.indexOf(":");
                if (split > -1) {
                    fieldName = msg.substring(2, split);
                    msg = msg.substring(split + 1);

                }
                result.add(new ValidateResult(fieldName, msg));
            }
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
        return result;
    }

    /**
     * 简单的excel转list
     *
     * @param filepath excel路径
     * @return 返回结果列表
     */
    public static List<Map<String, String>> excel2List(String filepath) {
        return excel2List(filepath, 0);
    }

    /**
     * 简单的excel转list
     *
     * @param filepath excel路径
     * @param  sheetIndex 工作表索引
     * @return 返回结果列表
     */
    public static List<Map<String, String>> excel2List(String filepath, int sheetIndex) {

        List<Map<String, String>> list = new ArrayList<>();
        FileInputStream inputStream = null;
        Workbook workbook;
        try {
            inputStream = new FileInputStream(filepath);
            workbook = StreamingReader.builder()
                    //缓存到内存中的行数，默认是10
                    .rowCacheSize(100)
                    //读取资源时，缓存到内存的字节大小，默认是1024
                    .bufferSize(4096)
                    //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
                    .open(inputStream);
        } catch (Exception e1) {
            if (inputStream == null) {
                throw new RuntimeException("无法读取文件");
            }
            try {
                workbook = new HSSFWorkbook(inputStream);
            } catch (Exception e2) {
                throw new RuntimeException(e2);
            }
        }

        Sheet sheet = workbook.getSheetAt(sheetIndex);

        Map<Integer, String> headMap = new HashMap<>();
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c);
                    if (null != cell) {
                        if (!cell.getStringCellValue().isEmpty()) {
                            headMap.put(c, cell.getStringCellValue());
                        }
                    }
                }
            } else {
                Map<String, String> obj = new HashMap<>();
                for (Integer i : headMap.keySet()) {
                    Cell cell = row.getCell(i);
                    //是否日期单元格
                    String dateFormat = "yyyy-MM-dd HH:mm:ss";
                    if (null != cell) {
                        String str;
                        CellType cellType = cell.getCellType();
                        //支持公式单元格
                        if (cellType == CellType.FORMULA) {
                            cellType = cell.getCachedFormulaResultType();
                        }
                        switch (cellType) {
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
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
                                str = null;
                                break;
                            case STRING:
                            default:
                                str = cell.getStringCellValue();
                                break;
                        }
                        obj.put(headMap.get(i), str);
                    }
                }
                list.add(obj);

            }
        }

        return list;
    }


    /**
     * 生成导出excel模板
     *
     * @param clazz 中间类类型
     * @param <T> 中间类类型
     * @return 导出excel模板
     */
    public static <T> byte[] createImportExcelTemplate(Class<T> clazz) {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Sheet1");
        int i = 0;
        XSSFRow row = sheet.createRow(0);
        for (Field field : clazz.getDeclaredFields()) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn != null) {
                XSSFCell cell = row.createCell(i);
                cell.setCellValue(excelColumn.value());
                i++;
            }
        }
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            wb.write(baos);
            baos.close();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
        return baos.toByteArray();
    }

    public static List<byte[]> getFloatImagesBytes(Sheet sheet, Integer rowIndex, Integer columnIndex) {
        List<byte[]> list = new ArrayList<>();
        for (Shape shape : sheet.getDrawingPatriarch()) {
            XSSFPicture picture = (XSSFPicture) shape;
            XSSFClientAnchor anchor = picture.getClientAnchor();
            if (anchor.getRow1() == rowIndex && anchor.getCol1() == columnIndex) {
                if (anchor.getRow1() != anchor.getRow2()) {
                    throw new ImageOutOfBoundsException();
                } else if (anchor.getCol1() != anchor.getCol2()) {
                    throw new ImageOutOfBoundsException();
                } else {
                    list.add(((XSSFPicture) shape).getPictureData().getData());
                }
            }
        }
        return list;
    }

    /**
     * 是否有图片字段
     *
     * @param clazz 中间类类型
     * @return 是否有图片字段
     */
    private static boolean hasCellImageField(Class<?> clazz) {
        for (Field field : clazz.getDeclaredFields()) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn != null) {
                if (field.getType().isAssignableFrom(BufferedImage.class)) {
                    return true;
                } else if (field.getType().isArray() && field.getType().getComponentType().equals(byte.class)) {
                    return true;
                }

            }
        }
        return false;
    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            String filepath, List<ColumnInfo> columnInfos,
            boolean faultTolerant) {
        return importExcel(filepath, columnInfos, faultTolerant, 1, null, null);
    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            String filepath, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            int startRow) {
        return importExcel(filepath, columnInfos, faultTolerant, startRow, null, null);
    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            String filepath, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            Function<Map<String, Object>, Boolean> importFunc) {
        return importExcel(filepath, columnInfos, faultTolerant, 1, null, importFunc);

    }

    /**
     * 导入excel
     *
     * @param filepath      excel文件路径
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            String filepath, List<ColumnInfo> columnInfos,
            boolean faultTolerant, int startRow,
            Function<Map<String, Object>, Boolean> importFunc) {
        return importExcel(filepath, columnInfos, faultTolerant, startRow, null, importFunc);
    }

    /**
     * 导入excel
     *
     * @param filepath           excel文件路径
     * @param columnInfos        列信息
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            String filepath, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            Function<Map<String, Object>, List<ValidateResult>> customValidateFunc,
            Function<Map<String, Object>, Boolean> importFunc) {
        return importExcel(filepath, columnInfos, faultTolerant, 1, customValidateFunc, importFunc);
    }

    /**
     * 导入excel
     *
     * @param filepath           excel文件路径
     * @param columnInfos        列信息
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow           开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            String filepath, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            int startRow,
            Function<Map<String, Object>, List<ValidateResult>> customValidateFunc,
            Function<Map<String, Object>, Boolean> importFunc) {
        FileInputStream inputStream;
        try {
            inputStream = new FileInputStream(filepath);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return importExcel(inputStream, columnInfos, faultTolerant, startRow, customValidateFunc, importFunc);
    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            InputStream inputStream, List<ColumnInfo> columnInfos,
            boolean faultTolerant) {
        return importExcel(inputStream, columnInfos, faultTolerant, 1, null, null);
    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult importExcel(
            InputStream inputStream, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            int startRow) {
        return importExcel(inputStream, columnInfos, faultTolerant, startRow, null, null);
    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            InputStream inputStream, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            Function<Map<String, Object>, Boolean> importFunc) {
        return importExcel(inputStream, columnInfos, faultTolerant, 1, null, importFunc);

    }

    /**
     * 导入excel
     *
     * @param inputStream   excel文件的字节数组
     * @param columnInfos   列信息
     * @param faultTolerant 是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow      开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param importFunc    一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            InputStream inputStream, List<ColumnInfo> columnInfos,
            boolean faultTolerant, int startRow,
            Function<Map<String, Object>, Boolean> importFunc) {
        return importExcel(inputStream, columnInfos, faultTolerant, startRow, null, importFunc);
    }

    /**
     * 导入excel
     *
     * @param inputStream        excel文件的字节数组
     * @param columnInfos        列信息
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            InputStream inputStream, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            Function<Map<String, Object>, List<ValidateResult>> customValidateFunc,
            Function<Map<String, Object>, Boolean> importFunc) {
        return importExcel(inputStream, columnInfos, faultTolerant, 1, customValidateFunc, importFunc);
    }

    /**
     * 导入excel
     *
     * @param inputStream        excel文件的字节数组
     * @param columnInfos        列信息
     * @param faultTolerant      是否容错，验证是所有数据先验证后在一条条导入。true表示不需要全部数据都符合验证，false则表示必须全部数据符合验证才执行导入。
     * @param startRow           开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
     * @param customValidateFunc {@code 自定义验证的方法，一般简单验证写在字段注解中，这里处理复杂验证，如身份证格式等，不需要请传null。如果验证错误,则返回List<ValidateResult>,由于一行数据可能有多个错误，所以用List。如果验证通过返回null或空list即可}
     * @param importFunc         一条条入库的方法,只有验证通过的数据才会进入此方法。如果你是批量入库，请自行获取结果的成功列表,此参数传null。返回true表示入库成功，入库失败提示请抛一个带message的Exception。
     * @return 导入结果
     */
    @Deprecated
    public static ImportResult<?> importExcel(
            InputStream inputStream, List<ColumnInfo> columnInfos,
            boolean faultTolerant,
            int startRow,
            Function<Map<String, Object>, List<ValidateResult>> customValidateFunc,
            Function<Map<String, Object>, Boolean> importFunc) {
        return importExcel(ExcelImportMapConfig.builder()
                .inputStream(inputStream)
                .columnInfos(columnInfos)
                .faultTolerant(faultTolerant)
                .startRow(startRow)
                .customValidateFunc(customValidateFunc)
                .importFunc(importFunc)
                .build());
    }

    public static ImportResult<?> importExcel(ExcelImportMapConfig config) {
        List<ColumnInfo> columnInfos = config.getColumnInfos();
        InputStream inputStream = config.getInputStream();
        if (inputStream == null && StringUtil.isNotBlank(config.getFilepath())) {
            try {
                inputStream = Files.newInputStream(Paths.get(config.getFilepath()));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        assert inputStream != null;
        int startRow = config.getStartRow();
        Function<Map<String, Object>, Boolean> importFunc = config.getImportFunc();
        Function<Map<String, Object>, List<ValidateResult>> customValidateFunc = config.getCustomValidateFunc();
        boolean faultTolerant = config.isFaultTolerant();
        ImportResult<Map<String, Object>> result = new ImportResult<>();
        result.setErrors(new ArrayList<>());
        Workbook workbook;
        try {
            if (columnInfos.stream().anyMatch(m -> m.getType() == ColumnType.IMAGE.getValue() || m.getType() == ColumnType.IMAGES.getValue())) {
                //如果有图片字段，则不使用StreamingWorkbook
                workbook = new XSSFWorkbook(inputStream);
            } else {
                workbook = StreamingReader.builder()
                        //缓存到内存中的行数，默认是10
                        .rowCacheSize(100)
                        //读取资源时，缓存到内存的字节大小，默认是1024
                        .bufferSize(4096)
                        //打开资源，必须，可以是InputStream或者是File，注意：只能打开XLSX格式的文件
                        .open(inputStream);
            }
        } catch (Exception e1) {
            try {
                workbook = new HSSFWorkbook(inputStream);
            } catch (Exception e2) {
                throw new RuntimeException(e2);
            }
        }
        Sheet sheet = workbook.getSheetAt(config.getSheetIndex());
        //列序号和字段的map
        Map<Integer, ColumnInfo> headMap = new HashMap<>();
        if (config.isOnlyData()) {
            columnInfos.stream().filter(m -> StringUtil.isNotEmpty(m.getColString())).forEach(columnInfo ->
            {
                int i = CellReference.convertColStringToIndex(columnInfo.getColString());
                headMap.put(i, columnInfo);
            });
        }
        Map<Integer, Map<String, Object>> list = new LinkedHashMap<>();
        int totalCount = 0;
        int titleRowNum = startRow - 1;
        for (Row row : sheet) {
            if (row.getRowNum() < titleRowNum) {
                //小于标题行的抛弃
            } else if (row.getRowNum() == titleRowNum && !config.isOnlyData()) {
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c);
                    if (null != cell) {
                        String title = cell.getStringCellValue();
                        ColumnInfo columnInfo = columnInfos.stream().filter(m -> StringUtil.isNotEmpty(m.getTitle()) && m.getTitle().equals(title)).findFirst().orElse(null);
                        if (columnInfo != null) {
                            headMap.put(c, columnInfo);
                        }
                    }
                }
                columnInfos.stream().filter(m -> StringUtil.isNotEmpty(m.getColString())).forEach(columnInfo ->
                {
                    int i = CellReference.convertColStringToIndex(columnInfo.getColString());
                    headMap.put(i, columnInfo);
                });

            } else {
                totalCount++;
                boolean isEmptyRow = true;
                Map<String, Object> obj = new HashMap<>();
                boolean validate = true;
                for (Integer c : headMap.keySet()) {
                    ColumnInfo columnInfo = headMap.get(c);
                    //处理图片
                    if (columnInfo.getType() == ColumnType.IMAGE.getValue()) {
                        List<byte[]> floatImages = getFloatImagesBytes(sheet, row.getRowNum(), c);
                        if (!CollectionUtils.isEmpty(floatImages)) {
                            obj.put(columnInfo.getName(), floatImages.get(0));
                            isEmptyRow = false;
                        }
                        continue;
                    } else if (columnInfo.getType() == ColumnType.IMAGES.getValue()) {
                        List<byte[]> floatImages = getFloatImagesBytes(sheet, row.getRowNum(), c);
                        obj.put(columnInfo.getName(), floatImages);
                        if (!CollectionUtils.isEmpty(floatImages)) {
                            isEmptyRow = false;
                        }
                        continue;
                    }
                    Cell cell = row.getCell(c);
                    //是否日期单元格
                    boolean isDateCell = SheetUtil.isDateCell(cell);
                    String dateFormat = "yyyy-MM-dd HH:mm:ss";
                    try {
                        if (null != cell) {
                            String str = SheetUtil.getCellStringValue(cell);
                            if (!StringUtil.isBlank(str)) {
                                isEmptyRow = false;
                            }
                            Object value = null;
                            if (isDateCell && (columnInfo.getType() == ColumnType.DATETIME.getValue() || columnInfo.getType() == ColumnType.DATE.getValue())) {
                                //特殊处理日期格式
                                if (!StringUtil.isBlank(str)) {
                                    value = StringUtil.parse(str, dateFormat, Date.class);
                                }
                            } else if (isDateCell && columnInfo.getType() == ColumnType.TIME.getValue()) {
                                if (!StringUtil.isBlank(str)) {
                                    value = StringUtil.parse(str, dateFormat, Time.class);
                                }
                            } else if (columnInfo.getType() == ColumnType.IMAGE.getValue()) {
                                value = SheetUtil.getCellImageBytes((XSSFWorkbook) workbook, cell);
                            } else if (columnInfo.getType() == ColumnType.LONG.getValue()) {
                                value = StringUtil.parse(str, Long.class);
                            } else if (columnInfo.getType() == ColumnType.DOUBLE.getValue()) {
                                value = StringUtil.parse(str, Double.class);
                            } else {
                                value = str;
                            }
                            obj.put(columnInfo.getName(), value);
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
                if (isEmptyRow) {
                    //空行时总数-1，且不校验
                    totalCount--;
                    continue;
                }
                List<ValidateResult> validateResults = ValidateUtil.validate(obj, columnInfos);
                validate = ValidateUtil.isValidate(result, headMap, row, validate, validateResults, columnInfos);

                if (customValidateFunc != null) {
                    List<ValidateResult> customValidateResults = customValidateFunc.apply(obj);
                    validate = ValidateUtil.isValidate(result, headMap, row, validate, customValidateResults, columnInfos);
                }
                if (validate) {
                    list.put(row.getRowNum(), obj);
                }
            }
        }
        //设置总记录数
        result.setTotalCount(totalCount);
        if (!list.isEmpty()) {
            if (faultTolerant || result.getErrors().isEmpty()) {
                //如果容错模式或是验证全部通过
                if (importFunc != null) {
                    //如果有导入函数
                    for (Map.Entry<Integer, Map<String, Object>> m : list.entrySet()) {
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


}
