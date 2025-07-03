# excelutil
基于java的excle工具类，主要是导入excel的前置校验工作。

# [按模板导出word](./doc/WordTemplateUtil.md)
## 一、当前最新版本
```xml
<dependency>
  <groupId>com.iceolive</groupId>
  <artifactId>excelutil</artifactId>
  <version>1.2.4</version>
</dependency>
```
## 二、快速开始
### 1.定义一个中间类，用于定义excel的列对象
验证注解基于validation-api包
```java
@Data
public class TestModel {
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
    // 图片，有两种方式嵌入，一种是浮动图片置于单元格内（不能越界），一种是嵌入单元格    
    // 单元格嵌入图片必须是图片公式为 =DISPIMG("ID_XXXX",1)，其中XXXX为32位十六进制字符串，只能有一张，类型只能是BufferedImage或byte[]
    // 浮动图片可以多张也可以单张，单张类型为BufferedImage或byte[],多张类型为 List<BufferedImage> 或 List<byte[]>（不支持数组是因为）
    // 只支持xlsx的单元格图片    
    @ExcelColumn("图片")
    private BufferedImage image;
}

```
### 2.调用

最简参数调用
```java
ImportResult importResult =  ExcelUtil.importExcel(ExcelImportConfig.<TestModel>builder()
        //中间类类型
        .clazz(TestModel.class)
        //excle文件路径, 支持xls和xlsx。
        .filepath(filepath)
        //是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
        .faultTolerant(true)
        .build()); 
```
全参数调用
```java
 ImportResult importResult =  ExcelUtil.importExcel(ExcelImportConfig.<TestModel>builder()
        //excle文件路径, 支持xls和xlsx。
        .filepath(filepath)
        //中间类类型
        .clazz(TestModel.class)
        //是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
        .faultTolerant(true)
        //开始行数，从1开始，当第一行是标题，则传1，当第二行是标题则传2。
        .startRow(1)
        //导出工作表索引
        .sheetIndex(0)
        //是否只有数据，当true时，startRow为数据开始行
        .onlyData(false)
        .customValidateFunc(
                m -> {
                //m为中间类对象
                //这里写自定义验证，比如身份证等用自定义注解无法验证的方法，不需要的话，此参数传null，或返回null或空list
                    List<ValidateResult> list = new ArrayList<>();
                    if (m.getName()!=null && !m.getName().startsWith("王")) {
                        //new ValidateResult第一个参数为字段名，框架根据字段名定位单元格地址
                        //new ValidateResult第二个参数为错误信息
                        list.add(new ValidateResult("name", "用户不姓王"));
                    }
                   return list; 
                })
        .importFunc(m -> {
                    //m为中间类对象
                    //这里写入库操作，返回true，则表示入库成功，如果入参出错，请抛异常，框架会捕获异常，错误信息为异常的getMessage()
                    boolean insertDBSuccess = yourInsertFunc(m);
                    if(insertDBSuccess){
                        return true;
                    }else{
                        throw new Exception("入库失败")
                    } 
                }));
```
 
### 3.返回结果
```java
    ImportResult importResult = ExcelUtil.importExcel(...);
    //导入成功的记录集，类型Map<Integer,T>,key为行号，value为中间类的对象
    importResult.getSucesses();
    //导入失败的记录集，类型List<ImportResult.ErrorMessage>
    //ImportResult.ErrorMessage 包括 row(行号),cell(单元格地址), message(错误信息)三个属性
    //ImportResult.ErrorMessage的行号是一定会有的，但是单元格地址在以下三种错误里面不会有。
    // 1.找不到列，由于没有列，所以也无法提示哪个单元格错误。
    // 2.自定义验证函数返回的ValidateResult写错了字段名，同1也是会提示找不到列
    // 3.入库函数时抛的异常也是不会提示到单元格。
    importResult.getErrors();
    //excel的总记录数，不包括标题
    importResult.getTotalCount();
    //导入成功的记录集，类型List<T>
    importResult.getSuccessList();
    //导入成功的记录条数。
    importResult.getSuccessCount();

```
### 4.json-schema验证
注解验证虽然方便易用。但如果同个实体存在不同验证规则的场景，写在代码上维护起来还是不太方便。所以增加了json-schema验证方法。    
注意：对于日期类型，json对应的是字符串，记得在实体类的字段上添加@JsonFormat注解       
4.1 定义json-schema
```json
{
  "$schema": "https://json-schema.org/draft/2019-09/schema",
  "type": "object",
  "properties": {
    "name": {
      "type": "string",
      "minLength": 2,
      "maxLength": 6
    },
    "age": {
      "type": "number"
    },
    "birth": {
      "type": [
        "string",
        "null"
      ],
      "format": "date"
    },
    "birth1": {
      "type": [
        "string",
        "null"
      ],
      "format": "date"
    },
    "agree": {
      "type": [
        "boolean",
        "null"
      ]
    }
  },
  "required": [
    "name",
    "age"
  ]
}
```
4.2 调用
```java
String schemaJson = yourLoadTextFromFile("schema.json")
 ImportResult importResult =  ExcelUtil.importExcel(ExcelImportConfig.<TestModel>builder()
        .filename("D://result.xlsx")
        .clazz(TestModel.class)
        .faultTolerant(true)
        .startRow(1)
        .customValidateFunc(m -> {
         return ExcelUtil.jsonSchemaValidate(schemaJson, m)
           return list; 
        }));
```
## 三、开发背景
项目起源是我想设计一个工具类，作为导入excel数据的通用处理工具。    
通常我们的excel模板并不是一一对应数据库的一张表。    
所以我这个工具也并不关心你的哪个字段对应数据库的哪个表。    
怎么入库是通过工具调用方(程序员)自己写委托方法自行处理。    
我关注的重点是导入失败时，是由于excel中的哪些内容导致的。    
这些错误信息最好是能一一反馈给最终的使用者(用户)。    
让用户知道导入失败时由于excel中的哪些单元格的什么原因导致的，以及已经导入成功的记录。    
灵活使用这个工具返回的信息。当导入失败时，是可以不需要程序员介入，而是由用户自行编辑修改模板的错误内容，并完成导入操作的。 
## 四、个人觉得的项目亮点
### 1.沿用spring的验证validation-api，减低学习成本
### 2.支持直接写自定义验证方法，方便扩展。
### 3.与数据库无关，你可以自己实现自己的持久化操作。
### 4.支持图片导入
