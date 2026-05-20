import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;
import com.iceolive.util.*;
import com.iceolive.util.annotation.ExcelColumn;
import com.iceolive.util.constants.ValidationConsts;
import com.iceolive.util.enums.ColumnType;
import com.iceolive.util.model.*;
import lombok.Data;
import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import javax.imageio.ImageIO;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import java.awt.image.BufferedImage;
import java.io.*;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

/**
 * ExcelUtil 综合测试类
 * <p>
 * 测试数据依赖 testdata 目录下的文件，运行前请确保文件存在。
 * 每个测试方法都会生成独立输出文件，方便人工检查。
 * </p>
 */
public class Tests {

    // region ---------- 测试模型 ----------

    @Data
    public static class TestModel {
        @NotNull
        @ExcelColumn("年龄")
        private Integer age;

        @NotBlank
        @ExcelColumn("姓名")
        private String name;

        @ExcelColumn(trueString = "是", falseString = "否")
        private Boolean agree;

        @ExcelColumn
        @JsonFormat(pattern = "yyyy-MM-dd")
        private Date birth;

        @ExcelColumn("birth")
        @JsonSerialize(using = LocalDateTimeSerializer.class)
        @JsonFormat(pattern = "yyyy-MM-dd")
        private LocalDateTime birth1;

        @ExcelColumn("time")
        @JsonFormat(pattern = "HH:mm:ss")
        private LocalTime time;

        @ExcelColumn("图片")
        private BufferedImage image;

        @ExcelColumn("图片2")
        private List<byte[]> image2;
    }

    // endregion

    // region ---------- 工具方法 ----------

    private static final String TESTDATA_DIR = System.getProperty("user.dir") + "//testdata//templates//";
    private static final String OUTPUT_DIR = System.getProperty("user.dir") + "//testdata//output//";

    static {
        new File(OUTPUT_DIR).mkdirs();
    }

    private String path(String filename) {
        return TESTDATA_DIR + filename;
    }

    private String outputPath(String filename) {
        return OUTPUT_DIR + filename;
    }

    private BufferedImage loadImage(String filename) throws IOException {
        return ImageIO.read(new File(path(filename)));
    }

    private void save(byte[] bytes, String filename) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(outputPath(filename))) {
            fos.write(bytes);
        }
        System.out.println("输出文件: " + path(filename));
    }

    // endregion

    // region ---------- Excel 导入测试 ----------

    /**
     * 测试1：注解方式导入 Excel（容错模式）
     */
    @Test
    public void test1_annotationImport() {
        ImportResult<TestModel> result = ExcelUtil.importExcel(
                ExcelImportConfig.<TestModel>builder()
                        .clazz(TestModel.class)
                        .filepath(path("test1.xlsx"))
                        .faultTolerant(true)
                        .build());
        System.out.println(result);
    }

    /**
     * 测试2：将 Excel 转为 List<Map>
     */
    @Test
    public void test2_excelToList() {
        List<Map<String, String>> list = ExcelUtil.excel2List(path("test1.xlsx"));
        System.out.println(list);
    }

    /**
     * 测试3：Map 方式导入（含规则校验）
     */
    @Test
    public void test3_mapImportWithRules() {
        List<ColumnInfo> columns = Arrays.asList(
                newColumn("age", "年龄", null, ColumnType.LONG,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED),
                        ColumnInfo.Rule.fromRange(1D, 99D, "年龄必须是1到99")),
                newColumn("name", "姓名", null, ColumnType.STRING,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("birth", "birth", null, ColumnType.DATETIME,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("agree", "agree", null, ColumnType.STRING,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("image", "图片", null, ColumnType.IMAGE,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("image2", "图片2", null, ColumnType.IMAGES,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED))
        );

        ImportResult<?> result = ExcelUtil.importExcel(
                ExcelImportMapConfig.builder()
                        .filepath(path("test1.xlsx"))
                        .faultTolerant(true)
                        .columnInfos(columns)
                        .build());
        System.out.println(result);
    }

    /**
     * 测试4：Map 方式导入（含正则、范围、枚举校验）
     */
    @Test
    public void test4_mapImportAdvancedRules() {
        List<ColumnInfo> columns = Arrays.asList(
                newColumn("age", null, "B", ColumnType.STRING,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED),
                        ColumnInfo.Rule.fromRegExp("\\d+", "年龄必须是数字"),
                        ColumnInfo.Rule.fromRange("1", "99", "年龄必须是1到99"),
                        ColumnInfo.Rule.fromEnums(Arrays.asList("4", "5", "99"), "年龄不在枚举范围")),
                newColumn("name", null, "A", ColumnType.STRING,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("birth", null, "C", ColumnType.DATETIME,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("agree", null, "D", ColumnType.STRING,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("image", null, "G", ColumnType.IMAGE,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED)),
                newColumn("image2", null, "F", ColumnType.IMAGES,
                        ColumnInfo.Rule.fromBuiltIn(ValidationConsts.REQUIRED))
        );

        ImportResult<?> result = ExcelUtil.importExcel(
                ExcelImportMapConfig.builder()
                        .filepath(path("test1.xlsx"))
                        .columnInfos(columns)
                        .faultTolerant(true)
                        .build());
        System.out.println(result);
    }

    // endregion

    // region ---------- Excel 导出测试 ----------

    /**
     * 测试5：导出 Excel（含图片）
     */
    @Test
    public void test5_exportExcelWithImages() throws IOException {
        List<Map<String, Object>> data = new ArrayList<>();

        Map<String, Object> item1 = new HashMap<>();
        item1.put("title", "标题1");
        item1.put("images", Collections.singletonList(
                ImageUtil.Image2Bytes(loadImage("20230627153447277.png"), "png")));
        data.add(item1);

        Map<String, Object> item2 = new HashMap<>();
        item2.put("title", "标题2");
        List<byte[]> images = new ArrayList<>();
        images.add(ImageUtil.Image2Bytes(loadImage("20230627153447823.png"), "png"));
        images.add(ImageUtil.Image2Bytes(loadImage("20230627153447850.png"), "png"));
        item2.put("images", images);
        data.add(item2);

        List<ColumnInfo> columns = Arrays.asList(
                new ColumnInfo("title", "标题", "A", ColumnType.STRING.getValue()),
                new ColumnInfo("images", "图片", "B", ColumnType.IMAGES.getValue()));

        try (FileInputStream fis = new FileInputStream(path("tpl.xlsx"))) {
            byte[] bytes = ExcelExportUtil.exportExcel(fis, data, columns, 1, true);
            save(bytes, "result_export.xlsx");
        }
    }

    /**
     * 测试6：导入 test5 生成的 Excel
     */
    @Test
    public void test6_importExportedExcel() {
        List<ColumnInfo> columns = Arrays.asList(
                new ColumnInfo("title", "标题", "A", ColumnType.STRING.getValue()),
                new ColumnInfo("images", "图片", "B", ColumnType.IMAGES.getValue()));

        ImportResult<?> result = ExcelUtil.importExcel(
                ExcelImportMapConfig.builder()
                        .filepath(outputPath("result_export.xlsx"))
                        .faultTolerant(true)
                        .columnInfos(columns)
                        .startRow(1)
                        .build());
        System.out.println(result);
    }

    // endregion

    // region ---------- 单表导入测试 ----------

    /**
     * 测试7：单表导入（固定单元格位置）
     */
    @Test
    public void test7_singleSheetImport() {
        List<FieldInfo> fields = Arrays.asList(
                newField("概况描述", "B13", ColumnType.STRING,
                        FieldInfo.Rule.fromRegExp("^.{6}$", "概况描述必须写6位")),
                new FieldInfo("产生原因", "E13", ColumnType.STRING.getValue()),
                new FieldInfo("涉及人员", "F13", ColumnType.STRING.getValue()),
                new FieldInfo("备注", "G13", ColumnType.STRING.getValue()));

        ImportSingleResult result = ExcelSingleUtil.importExcel(path("test2.xlsx"), fields);
        System.out.println(result);
    }

    // endregion

    // region ---------- Word 模板测试（基础） ----------

    /**
     * 测试8：Word 模板填充（基础示例）
     */
    @Test
    public void test8_wordTemplateBasic() throws IOException {
        StopWatch sw = new StopWatch();
        sw.start();

        List<Map<String, Object>> courses = new ArrayList<>();
        courses.add(newHashMap("name", "语文", "score", "99", "image", loadImage("20230627153447277.png")));
        courses.add(newHashMap("name", "数学", "score", "100", "image", loadImage("20230627153447823.png")));

        Map<String, Object> data = new HashMap<>();
        data.put("name", "张三");
        data.put("age", "20");
        data.put("desc", "换行\n换行\n换行");
        data.put("course", courses);
        data.put("image", loadImage("20230627153447850.png"));

        XWPFDocument doc = WordTemplateUtil.load(path("wordtpl.docx"));
        WordTemplateUtil.fillData(doc, data);
        WordTemplateUtil.save(doc, outputPath("result_basic.docx"));

        sw.stop();
        System.out.println("耗时: " + sw.getTime() + "ms");
    }

    // endregion

    // region ---------- Word 模板测试（新功能：图片尺寸） ----------

    /**
     * 测试9：Word 模板 - 图片原始尺寸（不传尺寸）
     * 模板占位符: @{image}
     */
    @Test
    public void test9_wordImageOriginalSize() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("image", loadImage("20260520140333960.png"));

        XWPFDocument doc = WordTemplateUtil.load(path("wordtpl_original.docx"));
        WordTemplateUtil.fillData(doc, data);
        WordTemplateUtil.save(doc, outputPath("result_img_original.docx"));
    }

    /**
     * 测试10：Word 模板 - 图片指定宽高
     * 模板占位符: @{image:200*150}
     */
    @Test
    public void test10_wordImageFixedSize() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("image", loadImage("20260520140333960.png"));

        XWPFDocument doc = WordTemplateUtil.load(path("wordtpl_fixed.docx"));
        WordTemplateUtil.fillData(doc, data);
        WordTemplateUtil.save(doc, outputPath("result_img_fixed.docx"));
    }

    /**
     * 测试11：Word 模板 - 图片只指定宽度（高度自动按比例）
     * 模板占位符: @{image:300*}
     */
    @Test
    public void test11_wordImageWidthOnly() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("image", loadImage("20260520140333960.png"));

        XWPFDocument doc = WordTemplateUtil.load(path("wordtpl_width.docx"));
        WordTemplateUtil.fillData(doc, data);
        WordTemplateUtil.save(doc, outputPath("result_img_width.docx"));
    }

    /**
     * 测试12：Word 模板 - 图片只指定高度（宽度自动按比例）
     * 模板占位符: @{image:*200}
     */
    @Test
    public void test12_wordImageHeightOnly() throws IOException {
        Map<String, Object> data = new HashMap<>();
        data.put("image", loadImage("20260520140333960.png"));

        XWPFDocument doc = WordTemplateUtil.load(path("wordtpl_height.docx"));
        WordTemplateUtil.fillData(doc, data);
        WordTemplateUtil.save(doc, outputPath("result_img_height.docx"));
    }

    // endregion

    // region ---------- Word 模板测试（业务场景） ----------

    /**
     * 测试13：Word 模板 - 受理告知单（业务场景）
     */
    @Test
    public void test13_wordTemplateBusiness() throws IOException {
        StopWatch sw = new StopWatch();
        sw.start();

        Map<String, Object> data = new HashMap<>();
        data.put("name", "张三");
        data.put("eventTitle", "邻里纠纷");
        data.put("year", "2026");
        data.put("date", "2026年1月1日");
        data.put("offices", "主办单位A、主办单位B、协办单位C");
        data.put("code", "01-291332142");
        data.put("mpName", "我的小程序");
        data.put("district", null);
        data.put("qrcode", loadImage("20230627153447277.png"));

        XWPFDocument doc = WordTemplateUtil.load(path("shouligao.docx"));
        WordTemplateUtil.fillData(doc, data);
        WordTemplateUtil.save(doc, outputPath("result_business.docx"));

        sw.stop();
        System.out.println("耗时: " + sw.getTime() + "ms");
    }

    // endregion

    // region ---------- 辅助方法 ----------

    private ColumnInfo newColumn(String field, String title, String col, ColumnType type, ColumnInfo.Rule... rules) {
        ColumnInfo ci = new ColumnInfo(field, title, col, type.getValue());
        if (rules.length > 0) {
            ci.setRules(new ArrayList<>(Arrays.asList(rules)));
        }
        return ci;
    }

    private FieldInfo newField(String name, String cell, ColumnType type, FieldInfo.Rule... rules) {
        FieldInfo fi = new FieldInfo(name, cell, type.getValue());
        if (rules.length > 0) {
            fi.setRules(new ArrayList<>(Arrays.asList(rules)));
        }
        return fi;
    }

    @SafeVarargs
    private final Map<String, Object> newHashMap(Object... kvs) {
        Map<String, Object> map = new HashMap<>();
        for (int i = 0; i < kvs.length; i += 2) {
            map.put((String) kvs[i], kvs[i + 1]);
        }
        return map;
    }

    // endregion
}
