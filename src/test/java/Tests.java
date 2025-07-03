import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;
import com.iceolive.util.*;
import com.iceolive.util.annotation.ExcelColumn;
import com.iceolive.util.constants.ValidationConsts;
import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.model.*;
import lombok.Data;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import javax.imageio.ImageIO;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import java.awt.image.BufferedImage;
import java.io.*;
import java.sql.Time;
import java.time.LocalDateTime;
import java.time.LocalTime;
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
        @ExcelColumn("time")
        @JsonFormat(pattern = "HH:mm:ss")
        private LocalTime time;

        @ExcelColumn("图片")
        private BufferedImage image;
        @ExcelColumn("图片2")
        private List<byte[]> image2;
    }

    @Test
    public void test1() {
        String filepath = System.getProperty("user.dir") + "//testdata//test1.xlsx";

        ImportResult<TestModel> importResult = ExcelUtil.importExcel(ExcelImportConfig.<TestModel>builder()
                //中间类类型
                .clazz(TestModel.class)
                //excle文件路径, 支持xls和xlsx。
                .filepath(filepath)
                //是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
                .faultTolerant(true)
                .build());
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
        ColumnInfo c1 = new ColumnInfo("age", "年龄", null, ColumnType.LONG.getValue());
        c1.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
                add(ColumnInfo.Rule.fromRange(1D, 99D, "年龄必须是1到99"));
            }
        });
        ColumnInfo c2 = new ColumnInfo("name", "姓名", null, ColumnType.STRING.getValue());
        c2.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c3 = new ColumnInfo("birth", "birth", null, ColumnType.DATETIME.getValue());
        c3.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c4 = new ColumnInfo("agree", "agree", null, ColumnType.STRING.getValue());
        c4.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c5 = new ColumnInfo("image", "图片", null, ColumnType.IMAGE.getValue());
        c5.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c6 = new ColumnInfo("image2", "图片2", null, ColumnType.IMAGES.getValue());
        c6.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        List<ColumnInfo> columnInfos = new ArrayList<>(Arrays.asList(c1, c2, c3, c4, c5, c6));
        ImportResult<?> importResult = ExcelUtil.importExcel(
                ExcelImportMapConfig.builder()
                        //excle文件路径,支持xls和xlsx。
                        .filepath(filepath)
                        //是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
                        .faultTolerant(true)
                        //列配置
                        .columnInfos(columnInfos)
                        .build());
        System.out.println(importResult);
    }

    @Test
    public void test4() {
        String filepath = System.getProperty("user.dir") + "//testdata//test1.xlsx";
        ColumnInfo c1 = new ColumnInfo("age", null, "B", ColumnType.STRING.getValue());
        c1.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
                add(ColumnInfo.Rule.fromRegExp("\\d+", "年龄必须是数字"));
                add(ColumnInfo.Rule.fromRange("1", "99", "年龄必须是1到99"));
                add(ColumnInfo.Rule.fromEnums(Arrays.asList("4", "5", "99"), "年龄不在枚举范围"));
            }
        });
        ColumnInfo c2 = new ColumnInfo("name", null, "A", ColumnType.STRING.getValue());
        c2.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c3 = new ColumnInfo("birth", null, "C", ColumnType.DATETIME.getValue());
        c3.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c4 = new ColumnInfo("agree", null, "D", ColumnType.STRING.getValue());
        c4.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c5 = new ColumnInfo("image", null, "G", ColumnType.IMAGE.getValue());
        c5.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c6 = new ColumnInfo("image2", null, "F", ColumnType.IMAGES.getValue());
        c6.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED));
            }
        });
        List<ColumnInfo> columnInfos = new ArrayList<>(Arrays.asList(c1, c2, c3, c4, c5, c6));
        ImportResult<?> importResult = ExcelUtil.importExcel(
                ExcelImportMapConfig.builder()
                        //excle文件路径,支持xls和xlsx。
                        .filepath(filepath)
                        //列配置
                        .columnInfos(columnInfos)
                        //是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
                        .faultTolerant(true)
                        .build());
        System.out.println(importResult);
    }

    @Test
    public void test5() throws IOException {
        String filepath = System.getProperty("user.dir") + "//testdata//tpl.xlsx";
        FileInputStream fileInputStream = new FileInputStream(filepath);
        List<Map<String, Object>> data = new ArrayList<>();
        Map<String, Object> item1 = new HashMap<>();
        item1.put("title", "标题1");
        List<byte[]> images = new ArrayList<byte[]>();

        BufferedImage bufferedImage = ImageIO.read(new File(System.getProperty("user.dir") + "//testdata//20230627153447277.png"));
        byte[] bytes = ImageUtil.Image2Bytes(bufferedImage, "png");
        images.add(bytes);

        item1.put("images", images);
        data.add(item1);
        Map<String, Object> item2 = new HashMap<>();
        item2.put("title", "标题2");
        images = new ArrayList<byte[]>();
        bufferedImage = ImageIO.read(new File(System.getProperty("user.dir") + "//testdata//20230627153447823.png"));
        bytes = ImageUtil.Image2Bytes(bufferedImage, "png");
        images.add(bytes);
        bufferedImage = ImageIO.read(new File(System.getProperty("user.dir") + "//testdata//20230627153447850.png"));
        bytes = ImageUtil.Image2Bytes(bufferedImage, "png");
        images.add(bytes);

        item2.put("images", images);
        data.add(item2);
        List<ColumnInfo> columnInfos = new ArrayList<>();
        columnInfos.add(new ColumnInfo("title", "标题", "A", ColumnType.STRING.getValue()));
        columnInfos.add(new ColumnInfo("images", "图片", "B", ColumnType.IMAGES.getValue()));
        bytes = ExcelExportUtil.exportExcel(fileInputStream, data, columnInfos, 1, true);
        String outputFile = System.getProperty("user.dir") + "//testdata//result.xlsx";
        FileOutputStream fos = new FileOutputStream(outputFile);
        fos.write(bytes);
        fos.close();
        System.out.println(outputFile);

    }

    @Test
    public void test6() {
        String filepath = System.getProperty("user.dir") + "//testdata//result.xlsx";
        List<ColumnInfo> columnInfos = new ArrayList<>();
        columnInfos.add(new ColumnInfo("title", "标题", "A", ColumnType.STRING.getValue()));
        columnInfos.add(new ColumnInfo("images", "图片", "B", ColumnType.IMAGES.getValue()));
        ImportResult<?> importResult = ExcelUtil.importExcel(
                ExcelImportMapConfig.builder()
                        .filepath(filepath)
                        .faultTolerant(true)
                        .columnInfos(columnInfos)
                        .startRow(1)
                        .build());
        System.out.println(importResult);
    }

    @Test
    public void test7() {
        String filepath = System.getProperty("user.dir") + "//testdata//test2.xlsx";
        List<FieldInfo> fieldInfos = new ArrayList<>();
        FieldInfo fieldInfo = new FieldInfo("概况描述", "B13", ColumnType.STRING.getValue());
        fieldInfo.setRules(new ArrayList<BaseInfo.Rule>() {{
            add(BaseInfo.Rule.fromRegExp("^.{6}$", "概况描述必须写6位"));
        }});
        fieldInfos.add(fieldInfo);
        fieldInfos.add(new FieldInfo("产生原因", "E13", ColumnType.STRING.getValue()));
        fieldInfos.add(new FieldInfo("涉及人员", "F13", ColumnType.STRING.getValue()));
        fieldInfos.add(new FieldInfo("备注", "G13", ColumnType.STRING.getValue()));
        ImportSingleResult importSingleResult = ExcelSingleUtil.importExcel(filepath, fieldInfos);
        System.out.println(importSingleResult);
    }

    @Test
    public void test8() throws IOException {
        String filepath = System.getProperty("user.dir") + "//testdata//wordtpl.docx";
        Map<String, Object> map = new HashMap<>();
        List<Map<String, Object>> list = new ArrayList<>();
        list.add(new HashMap<String, Object>() {{
            put("name", "语文");
            put("score", "99");
            put("image",ImageIO.read(new File(System.getProperty("user.dir") + "//testdata//20230627153447277.png")));
        }});
        list.add(new HashMap<String, Object>() {{
            put("name", "数学");
            put("score", "100");
            put("image",ImageIO.read(new File(System.getProperty("user.dir") + "//testdata//20230627153447823.png")));
        }});
        map.put("name", "张三");
        map.put("age", "20");
        map.put("desc", "换行\n换行\n换行");
        map.put("course", list);
        map.put("image",ImageIO.read(new File(System.getProperty("user.dir") + "//testdata//20230627153447850.png")));
        XWPFDocument doc = WordTemplateUtil.load(filepath);
        WordTemplateUtil.fillData(doc, map);
        WordTemplateUtil.save(doc, System.getProperty("user.dir") + "//testdata//result.docx");
    }

    @Test
    public void test9() throws IOException {

        String filepath = System.getProperty("user.dir") + "//testdata//tpl.xlsx";
        FileInputStream fileInputStream = new FileInputStream(filepath);
        List<ColumnInfo> columnInfos = new ArrayList<>();
        columnInfos.add(new ColumnInfo("title", "标题", "A", ColumnType.STRING.getValue()));
        columnInfos.add(new ColumnInfo("images", "图片", "B", ColumnType.IMAGES.getValue()));
        ColumnInfo c3 = new ColumnInfo("enums", "枚举", "C", ColumnType.STRING.getValue());
        c3.setRules(new ArrayList<BaseInfo.Rule>() {{
            add(BaseInfo.Rule.fromEnums(Arrays.asList("澄海区", "金平区", "龙湖区"), "枚举值错误"));
        }});
        columnInfos.add(c3);
        ColumnInfo c4 = new ColumnInfo("range", "范围", "D", ColumnType.LONG.getValue());
        c4.setRules(new ArrayList<BaseInfo.Rule>() {{
            add(BaseInfo.Rule.fromRange(1, 5, "范围错误"));
        }});
        columnInfos.add(c4);

        byte[] bytes = ExcelExportUtil.setDataValidationRules(fileInputStream, columnInfos, 1);
        String outputFile = System.getProperty("user.dir") + "//testdata//result.xlsx";
        FileOutputStream fos = new FileOutputStream(outputFile);
        fos.write(bytes);
        fos.close();
        System.out.println(outputFile);
    }
}
