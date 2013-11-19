package com.ecfront.easybi.excelconverter.test.tojava.project;


import com.ecfront.easybi.excelconverter.exchange.EZExcel;
import junit.framework.Assert;
import org.junit.Test;

public class ProjectTest {

    @Test
    public void test() throws Exception {
        Project project = EZExcel.toJava(ProjectTest.class.getResource("/").getFile() + "project.xlsx", Project.class);
        Assert.assertEquals(project.name, "XXX项目");
        Assert.assertEquals(project.totalPrice.toString(), "710.3");
        Assert.assertEquals(project.modules.get(3).level, 4);
    }
}
