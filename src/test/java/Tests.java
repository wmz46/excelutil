import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;
import com.iceolive.util.ExcelUtil;
import com.iceolive.util.annotation.ExcelColumn;
import com.iceolive.util.model.ImportResult;
import lombok.Data;
import org.junit.Test;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import java.awt.image.BufferedImage;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class Tests {
    @Data
    public static class TestModel {
        @NotNull //验证注解
        @ExcelColumn("年龄")//注解excel的列标题
        private Integer age;
        @NotBlank//验证注解
        @ExcelColumn("姓名")//注解excel的列标题
        private String name;
        @ExcelColumn(trueString = "是",falseString = "否") //支持自定义布尔值
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
    public void test1(){
        String filepath = System.getProperty("user.dir")+ "//testdata//test1.xlsx";
        ImportResult importResult =  ExcelUtil.importExcel(filepath,//excle文件路径,也传excle文件的字节数组byte[],支持xls和xlsx。
                TestModel.class,//中间类类型
                true//是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
        );
        System.out.println(importResult);
    }
    @Test
    public void test2(){
        String filepath = System.getProperty("user.dir")+ "//testdata//test1.xlsx";
        List<Map<String, String>> list = ExcelUtil.excel2List(filepath);
        System.out.println(list);
    }
}
