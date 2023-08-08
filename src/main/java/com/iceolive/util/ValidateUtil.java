package com.iceolive.util;

import com.iceolive.util.constants.ValidationConsts;
import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.enums.RuleType;
import com.iceolive.util.model.BaseInfo;
import com.iceolive.util.model.ColumnInfo;
import com.iceolive.util.model.ImportResult;
import com.iceolive.util.model.ValidateResult;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

public class ValidateUtil {
    public static List<ValidateResult> validate(Map<String, Object> obj, List<? extends BaseInfo> columnInfos) {
        List<ValidateResult> result = new ArrayList<>();
        for (BaseInfo columnInfo : columnInfos) {
            String name = columnInfo.getName();
            if (!CollectionUtils.isEmpty(columnInfo.getRules())) {
                Object value = obj.get(name);
                for (ColumnInfo.Rule rule : columnInfo.getRules()) {
                    String code = rule.getCode();
                    String msg = rule.getMessage();
                    RuleType ruleType = rule.getType();
                    if (value != null) {
                        if (!Arrays.asList(ColumnType.IMAGE, ColumnType.IMAGES).contains(columnInfo.getType())) {
                            if (Arrays.asList(RuleType.BUILTIN, RuleType.REGEXP).contains(ruleType)) {
                                String regex = null;
                                if (ruleType == RuleType.REGEXP) {
                                    //正则
                                    regex = code;
                                    if (StringUtil.isEmpty(msg)) {
                                        msg = "参数输入有误";
                                    }
                                } else if (ValidationConsts.EMAIL.equals(code)) {
                                    regex = "^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+$";
                                    if (StringUtil.isEmpty(msg)) {
                                        msg = "请输入正确的邮箱地址";
                                    }
                                } else if (ValidationConsts.MOBILE.equals(code)) {
                                    regex = "^1[0-9]{10}$";
                                    if (StringUtil.isEmpty(msg)) {
                                        msg = "请输入正确的手机号";
                                    }
                                } else if (ValidationConsts.REQUIRED.equals(code)) {
                                    regex = "^[\\s\\S]+$";
                                    if (StringUtil.isEmpty(msg)) {
                                        msg = "参数不能为空";
                                    }
                                }
                                if (!StringUtil.isEmpty(regex)) {
                                    if (!Pattern.matches(regex, String.valueOf(value))) {
                                        result.add(new ValidateResult(name, msg));
                                    }
                                } else if (ValidationConsts.IDCARD.equals(code)) {
                                    if (StringUtil.isEmpty(msg)) {
                                        msg = "请输入正确的身份证号";
                                    }
                                    if (!IdCardUtil.validate(String.valueOf(value))) {
                                        result.add(new ValidateResult(name, msg));
                                    }
                                }
                            } else if (RuleType.ENUMS == ruleType) {
                                //枚举校验
                                if (rule.getEnumValues() == null || !rule.getEnumValues().contains(String.valueOf(value))) {
                                    result.add(new ValidateResult(name, msg));
                                }
                            } else if (RuleType.RANGE == ruleType) {
                                //范围校验
                                try {
                                    if (!NumberUtil.lessOrEqual(rule.getMin(), value)) {
                                        result.add(new ValidateResult(name, msg));
                                    } else if (!NumberUtil.greaterOrEqual(rule.getMax(), value)) {
                                        result.add(new ValidateResult(name, msg));
                                    }
                                }catch (Exception e){
                                    result.add(new ValidateResult(name,msg));
                                }
                            }
                        } else {
                            if (value.getClass().isAssignableFrom(ArrayList.class)) {
                                if (ValidationConsts.REQUIRED.equals(code) && CollectionUtils.isEmpty((List) value)) {
                                    if (StringUtil.isEmpty(msg)) {
                                        msg = "参数不能为空";
                                    }
                                    result.add(new ValidateResult(name, msg));
                                }
                            }
                        }
                    } else {
                        if (ValidationConsts.REQUIRED.equals(code)) {
                            if (StringUtil.isEmpty(msg)) {
                                msg = "参数不能为空";
                            }
                            result.add(new ValidateResult(name, msg));
                        }
                    }

                }
            }
        }
        return result;
    }

    public static boolean isValidate(ImportResult result, Map<Integer, ColumnInfo> headMap, Row row, boolean validate, List<ValidateResult> validateResults, List<ColumnInfo> columnInfos) {
        if (validateResults != null && !validateResults.isEmpty()) {
            validate = false;
            for (ValidateResult v : validateResults) {
                //错误是否在单元格内
                boolean errorInCell = false;
                for (Map.Entry<Integer, ColumnInfo> m : headMap.entrySet()) {
                    ColumnInfo columnInfo = m.getValue();
                    if (columnInfo.getName().equals(v.getFieldName())) {
                        ImportResult.ErrorMessage errorMessage = new ImportResult.ErrorMessage();
                        errorMessage.setRow(row.getRowNum());
                        errorMessage.setCol(CellReference.convertNumToColString(m.getKey()));
                        errorMessage.setCell(new CellAddress(row.getRowNum(), m.getKey()).toString());
                        errorMessage.setMessage(v.getMessage());
                        result.getErrors().add(errorMessage);
                        errorInCell = true;
                        break;
                    }
                }
                if (!errorInCell) {
                    String fieldName = v.getFieldName();
                    String columnName = fieldName;
                    ColumnInfo columnInfo = columnInfos.stream().filter(m -> m.getName().equals(fieldName)).findFirst().orElse(null);
                    if (columnInfo != null) {
                        String title = columnInfo.getTitle();
                        if (StringUtil.isNotEmpty(title)) {
                            columnName = title;
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
}
