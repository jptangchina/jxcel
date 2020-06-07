[![](https://img.shields.io/badge/version-1.0.2--alpha1-orange.svg)](https://mvnrepository.com/artifact/com.jptangchina/jxcel/1.0.0 )
[![](https://img.shields.io/badge/license-MIT-brightgreen.svg)](https://mit-license.org )


## 说明（非常重要）
1. 之前错误的将开发版作为正式版本发布到了中央仓库（其实应该还是Alpha版，只是为了玩一玩才发布的），因此将最新版发布为Alpha版本重新发布
2. 如果不是特别追求客制化的需求，阿里的[easyexcel](https://github.com/alibaba/easyexcel)我是极为推荐的
3. 这个版本的代码是用的Apache poi操作的（存在内存溢出可能，可自行搜索引擎），所以在内存方面的问题是需要使用者考量的(后面有时间或许会基于easyexcel做二次开发)
4. 关于反射的效率问题，反射的结果应该需要做缓存保存起来而不是每次都执行。或者干脆换用其他的反射或者实现方式
5. 内测版还存在很多的bug，代码的水平也还很拙劣。如果只是作为一个想法上的参考我觉得是可以的，但是如果投入实际使用还请结合前面几点三思
6. 暂时还没啥想到要写的，欢迎提issue

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