package com.jptangchina.jxcel;

import com.jptangchina.jxcel.bean.Student;
import org.junit.BeforeClass;
import org.junit.Test;

import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class JxcelGenratorTest {

    private static JxcelGenrator xlsGenerator;
    private static JxcelGenrator xlsxGenerator;

    @BeforeClass
    public static void init() {
        xlsGenerator = JxcelGenrator.xlsGenrator();
        xlsxGenerator = JxcelGenrator.xlsxGenrator();
    }

    @Test
    public void testGenerateToWorkbook() {
        Student student = new Student(18, 0, "JptangChina", new Date(), "18510010000");
        List<Student> list = Arrays.asList(student);
        xlsxGenerator.generateFile(list, "/Users/jptang/Desktop/jxcel.xlsx");
    }

}
