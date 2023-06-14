import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;
import com.iceolive.util.ExcelUtil;
import com.iceolive.util.annotation.ExcelColumn;
import com.iceolive.util.constants.ValidationConsts;
import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.model.ColumnInfo;
import com.iceolive.util.model.ImportResult;
import lombok.Data;
import org.junit.Test;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import java.awt.image.BufferedImage;
import java.time.LocalDateTime;
import java.util.*;

public class Tests {
    @Data
    public static class TestModel {
        @NotNull //验证注解
        @ExcelColumn("年龄")//注解excel的列标题
        private Integer age;
        @NotBlank//验证注解
        @ExcelColumn("姓名")//注解excel的列标题
        private String name;
        @ExcelColumn(trueString = "是", falseString = "否") //支持自定义布尔值
        private Boolean agree;
        @ExcelColumn//支持日期类型
        @JsonFormat(pattern = "yyyy-MM-dd")//如果使用json-schema验证，必须添加
        private Date birth;
        @ExcelColumn("birth")//支持一列匹配多个属性
        @JsonSerialize(using = LocalDateTimeSerializer.class)//如果使用json-schema验证，必须添加
        @JsonFormat(pattern = "yyyy-MM-dd")//如果使用json-schema验证，必须添加
        private LocalDateTime birth1;

        @ExcelColumn("图片")
        private BufferedImage image;
        @ExcelColumn("图片2")
        private List<byte[]> image2;
    }

    @Test
    public void test1() {
        String filepath = System.getProperty("user.dir") + "//testdata//test1.xlsx";
        ImportResult importResult = ExcelUtil.importExcel(filepath,//excle文件路径,也传excle文件的字节数组byte[],支持xls和xlsx。
                TestModel.class,//中间类类型
                true//是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
        );
        System.out.println(importResult);
    }

    @Test
    public void test2() {
        String filepath = System.getProperty("user.dir") + "//testdata//test1.xlsx";
        List<Map<String, String>> list = ExcelUtil.excel2List(filepath);
        System.out.println(list);
    }

    @Test
    public void test3() {
        String filepath = System.getProperty("user.dir") + "//testdata//test1.xlsx";
        List<ColumnInfo> columnInfos = new ArrayList<>();
        ColumnInfo c1= new ColumnInfo("age","年龄",ColumnType.LONG);
        c1.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
                add(new ColumnInfo.Rule("/^(?:[1-9]|[1-9][0-9])$/","年龄必须是1到99"));
            }
        });
        ColumnInfo c2= new ColumnInfo("name","姓名",ColumnType.STRING);
        c2.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c3= new ColumnInfo("birth","birth",ColumnType.DATETIME);
        c3.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c4= new ColumnInfo("agree","agree",ColumnType.STRING);
        c4.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c5= new ColumnInfo("image","图片",ColumnType.IMAGE);
        c5.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c6= new ColumnInfo("image2","图片2",ColumnType.IMAGES);
        c6.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        columnInfos.addAll(Arrays.asList(c1,c2,c3,c4,c5,c6));
        ImportResult importResult = ExcelUtil.importExcel(filepath,//excle文件路径,也传excle文件的字节数组byte[],支持xls和xlsx。
                columnInfos,//中间类类型
                true//是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
        );
        System.out.println(importResult);
    }
}
