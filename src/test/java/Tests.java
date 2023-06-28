import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;
import com.iceolive.util.ExcelExportUtil;
import com.iceolive.util.ExcelUtil;
import com.iceolive.util.ImageUtil;
import com.iceolive.util.annotation.ExcelColumn;
import com.iceolive.util.constants.ValidationConsts;
import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.model.ColumnInfo;
import com.iceolive.util.model.ImportResult;
import lombok.Data;
import org.junit.Test;

import javax.imageio.ImageIO;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import java.awt.image.BufferedImage;
import java.io.*;
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
        ColumnInfo c1 = new ColumnInfo("age", "年龄", null, ColumnType.LONG.getValue());
        c1.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
                add(new ColumnInfo.Rule("/^(?:[1-9]|[1-9][0-9])$/", "年龄必须是1到99"));
            }
        });
        ColumnInfo c2 = new ColumnInfo("name", "姓名", null, ColumnType.STRING.getValue());
        c2.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c3 = new ColumnInfo("birth", "birth", null, ColumnType.DATETIME.getValue());
        c3.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c4 = new ColumnInfo("agree", "agree", null, ColumnType.STRING.getValue());
        c4.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c5 = new ColumnInfo("image", "图片", null, ColumnType.IMAGE.getValue());
        c5.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c6 = new ColumnInfo("image2", "图片2", null, ColumnType.IMAGES.getValue());
        c6.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        columnInfos.addAll(Arrays.asList(c1, c2, c3, c4, c5, c6));
        ImportResult importResult = ExcelUtil.importExcel(filepath,//excle文件路径,也传excle文件的字节数组byte[],支持xls和xlsx。
                columnInfos,//中间类类型
                true//是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
        );
        System.out.println(importResult);
    }

    @Test
    public void test4() {
        String filepath = System.getProperty("user.dir") + "//testdata//test1.xlsx";
        List<ColumnInfo> columnInfos = new ArrayList<>();
        ColumnInfo c1 = new ColumnInfo("age", null, "B", ColumnType.LONG.getValue());
        c1.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
                add(new ColumnInfo.Rule("/^(?:[1-9]|[1-9][0-9])$/", "年龄必须是1到99"));
            }
        });
        ColumnInfo c2 = new ColumnInfo("name", null, "A", ColumnType.STRING.getValue());
        c2.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c3 = new ColumnInfo("birth", null, "C", ColumnType.DATETIME.getValue());
        c3.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c4 = new ColumnInfo("agree", null, "D", ColumnType.STRING.getValue());
        c4.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c5 = new ColumnInfo("image", null, "G", ColumnType.IMAGE.getValue());
        c5.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        ColumnInfo c6 = new ColumnInfo("image2", null, "F", ColumnType.IMAGES.getValue());
        c6.setRules(new ArrayList<ColumnInfo.Rule>() {
            {
                add(new ColumnInfo.Rule(ValidationConsts.REQUIRED));
            }
        });
        columnInfos.addAll(Arrays.asList(c1, c2, c3, c4, c5, c6));
        ImportResult importResult = ExcelUtil.importExcel(filepath,//excle文件路径,也传excle文件的字节数组byte[],支持xls和xlsx。
                columnInfos,//中间类类型
                true//是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
        );
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
    public void test6(){
        String filepath = System.getProperty("user.dir") + "//testdata//result.xlsx";
        List<ColumnInfo> columnInfos = new ArrayList<>();
        columnInfos.add(new ColumnInfo("title", "标题", "A", ColumnType.STRING.getValue()));
        columnInfos.add(new ColumnInfo("images", "图片", "B", ColumnType.IMAGES.getValue()));
        ImportResult importResult = ExcelUtil.importExcel(filepath, columnInfos, true, 1);
        System.out.println(importResult);
    }
}
