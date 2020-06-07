package com.jptangchina.jxcel;

import com.jptangchina.jxcel.bean.Student;
import org.junit.Assert;
import org.junit.Test;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.util.List;

public class JxcelParserTest {

    @Test
    public void testParseFromFile() {
        String desktop = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
        List<Student> result = JxcelParser.parser().parseFromFile(Student.class, new File(desktop + "/jxcel.xlsx"));
        Assert.assertTrue(result.size() > 0);
    }

}
