# excelutil
基于java的excle工具类
## 一、当前最新版本
```xml
<dependency>
  <groupId>com.iceolive</groupId>
  <artifactId>excelutil</artifactId>
  <version>0.0.1</version>
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
}

```
### 2.调用
```java
 ImportResult importResult =  ExcelUtil.importExcel("D://result.xlsx",//excle文件路径,也传excle文件的字节数组byte[],支持xls和xlsx。
                 TestModel.class,//中间类类型
                true,//是否容错处理，false则全部数据验证必须通过才执行入库操作，且入库操作只要没返回true，则不继续执行。true则只会对验证成功的记录进行入库操作，入库操作失败不影响后面的入库。
                m -> {
                //m为中间类对象
                //这里写自定义验证，比如身份证等用自定义注解无法验证的方法，不需要的话，此参数传null，或返回null或空list
                    List<ValidateResult> list = new ArrayList<>();
                    if (!m.getName().startsWith("王")) {
                        list.add(new ValidateResult("name", "用户不姓王"));
                    }
                   return list; 
                }, m -> {
                    //m为中间类对象
                    //这里写入库操作，返回true，则表示入库成功，如果入参出错，请抛异常，框架会捕获异常，错误信息为异常的getMessage()
                    boolean insertDBSuccess = yourInsertFunc(m);
                    if(insertDBSuccess){
                        return true;
                    }else{
                        throw new Exception("入库失败")
                    } 
                });
```
### 3.返回结果
```java
    ImportResult importResult = ExcelUtil.importExcel(...);
    importResult.getSucesses();//导入成功的记录集，类型Map<Integer,T>,key为行号，value为中间类的对象
    importResult.getErrors();//导入失败的记录集，类型List<ErrorMessage>，ErrorMessage 包括 row(行号),cell(单元格地址), message(错误信息)三个属性
    importResult.getTotalCount();//excel的总记录数，不包括标题
    importResult.getSuccessList();//导入成功的记录集，类型List<T>
    importResult.getSuccessCount();//导入成功的记录条数。

```
## 三、开发背景
项目起源是我想设计一个工具类，作为导入excel数据的通用处理工具。    
通常我们的excel模板并不是一一对应数据库的一张表。    
所以我这个工具也并不关心你的哪个字段对应数据库的哪个表。    
怎么入库是通过工具调用方(程序员)自己写委托方法自行处理。    
我关注的重点是导入失败时，是由于excel中的哪些内容导致的。    
这些错误信息最好是能一一反馈给最终的使用者(用户)。    
让用户知道导入失败时由于excel中的哪些单元格的什么原因导致的，以及已经导入成功的记录。    
灵活使用这个工具返回的信息。当导入失败时，是可以不需要程序员介入，而是由用户自行编辑修改，并完成导入操作的。 