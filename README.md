[![](https://img.shields.io/badge/version-1.0.1-orange.svg)](https://mvnrepository.com/artifact/com.jptangchina/jxcel/1.0.0 )
[![](https://img.shields.io/badge/license-MIT-brightgreen.svg)](https://mit-license.org )


## Jxcel简介
Jxcel是一个支持Java对象与Excel（目前仅xlsx、xls）互相转换的工具包。

## 特性说明
* Java对象输出为Excel文件或Workbook对象
* 语义化转换，将数字类型或布尔类型的值与语义化的值互相转换
* 生成的Excel文件可以对列进行排序
* 表头与Java属性精确匹配
* 支持几乎所有基本数据类型以及日期类型的转换
* 日期格式自定义
* 表格宽度自适应
* ......更多特性

## 引入依赖包
以Maven为例，引入Jxcel依赖包：
```xml
<dependency>
    <groupId>com.jptangchina</groupId>
    <artifactId>jxcel</artifactId>
    <version>${jxcel.version}</version>
</dependency>
```
## 准备数据模型
```java
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@JxcelSheet("学生名单")
class Student {
    @JxcelCell("年龄")
    private int age;
    @JxcelCell(value = "性别", parse = {"男", "女"})
    private int sex;
    @JxcelCell(value = "姓名", order = 1)
    private String name;
    @JxcelCell(value = "出生日期", format = "yyyy-MM-dd")
    private Date birthDay;
    @JxcelCell(value = "手机号", suffix = "\t")
    private String mobile;
}
```
## 导出数据到Excel
```java
// 导出为XLS Workbook对象
JxcelGenrator.xlsGenrator().generateWorkbook(Arrays.asList(new Student()));
// 导出为XLSX Workbook对象
JxcelGenrator.xlsxGenrator().generateWorkbook(Arrays.asList(new Student()));
// 导出为XLS文件
JxcelGenrator.xlsGenrator().generateFile(Arrays.asList(new Student()));
// 导出为XLSX文件
JxcelGenrator.xlsxGenrator().generateFile(Arrays.asList(new Student()));
```
## 将Excel解析为Java对象
```java
// 从文件解析
JxcelParser.parser().parseFromFile(Student.class, new File(filePath));
// 从Workbood对象解析
JxcelParser.parser().parseFromWorkbook(Student.class, workbook);
```
## 例子
```java
Student student = new Student(18, 0, "JptangChina", new Date(), "18510010000");
JxcelGenrator.xlsxGenrator().generateFile(Arrays.asList(student), "/home/jptangchina/test.xlsx");
```
输出的表格如下：

![](https://s2.ax1x.com/2019/06/14/V4Zy8J.jpg)

欢迎来我的个人博客参观指导：https://www.jptangchina.com