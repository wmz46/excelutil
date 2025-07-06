# WordTemplateUtil

根据模板导出word，支持图片导出

## 用法

```java
// 加载word
XWPFDocument doc = WordTemplateUtil.load(filepath);
//填充数据
WordTemplateUtil.fillData(doc, map);
//保存word
WordTemplateUtil.save(doc, System.getProperty("user.dir") + "//testdata//result.docx");
```
示例数据

```json
{
  "description": "采用Base64图片数据的JSON示例，适用于文档导入",
  "main_image": {
    "name": "示例主图",
    "type": "image/png",
    "data": "iVBORw0KGgoAAAANSUhEUgAA...（完整Base64数据）"
  },
  "items": [
    {
      "name": "列表项1",
      "image_base64": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAA...（缩略）"
    },
    {
      "name": "列表项2",
      "image_base64": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAA...（缩略）"
    }
  ]
}
```
模板占位符
### 普通变量
```txt
${description}  ${main_image.name}
```
### 图片
图片类型：支持BufferImage | byte[] | base64(String) | Data Url (String)（本质就是base64加了data:image/png;base64,前缀）
 不支持网络路径，不想在框架中额外请求外部地址，避免下载木马文件或额外鉴权处理
```txt
@{main_image_base64:100*200} //表示宽100，高200的图片 
@{main_image.data:100*200}
@{main_image.data} //不指定尺寸，将按原图尺寸输出
```
### 表格单行循环
使用`列表变量[].#index`获取当前循环索引    
**不支持多行循环**
```txt
| 行号 | 字符串           | 图片                             |
| ${item[].#index+1}  | ${items[].name} | @{items[].image_base64:100*200} |
```

### 表格单列循环
使用`列表变量[].#index`获取当前循环索引    
**不支持多列循环**
```txt
|  序号 |  ${item[].#index+1#col}                    |
| 字符串 |  ${items[].name#col}            |
| 图片   | @{items[].image_base64:100*200#col} |
```
### 单段落循环
使用`列表变量[].#index`获取当前循环索引    
**不支持多段落循环，如需多段落循环，可通过单列无边框表格或块级循环实现。**
```txt
${#index}   ${items[].name}   @{items[].image_base64:100*200}
```
### 块级条件（开发中）
```txt
{{#if name == '张三'}}

{{/if}}

```
### 块级循环（开发中）
```txt
{{for item,i in items}}
{{i}} {{item.name}}
{{/for}}
或
{{for item in items}}
{{item.name}}
{{/for}}
```