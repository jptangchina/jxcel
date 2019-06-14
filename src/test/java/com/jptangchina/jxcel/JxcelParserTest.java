package com.jptangchina.jxcel;

import com.jptangchina.jxcel.bean.Student;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class JxcelParserTest {

    @Test
    public void testParseFromFile() {
        List<Student> result = JxcelParser.parser().parseFromFile(Student.class, new File("/Users/jptang/Desktop/jxcel.xlsx"));
        Assert.assertTrue(result.size() > 0);
    }

}
