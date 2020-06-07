package com.jptangchina.jxcel;

import com.jptangchina.jxcel.bean.Student;
import org.junit.BeforeClass;
import org.junit.Test;

import javax.swing.filechooser.FileSystemView;
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
        Student student2 = new Student(16, 1, "JptangChina2", null, "12345");
        List<Student> list = Arrays.asList(student, student2);
        String desktop = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
        xlsxGenerator.generateFile(list, desktop + "/jxcel.xlsx");
    }

}
