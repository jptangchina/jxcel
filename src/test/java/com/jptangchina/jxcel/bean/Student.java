package com.jptangchina.jxcel.bean;

import com.jptangchina.jxcel.annotation.JxcelCell;
import com.jptangchina.jxcel.annotation.JxcelSheet;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Date;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@JxcelSheet("学生名单")
public class Student {

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
