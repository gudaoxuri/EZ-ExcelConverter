package com.ecfront.easybi.excelconverter.test.tojava.project;


import com.ecfront.easybi.excelconverter.exchange.EZExcel;
import junit.framework.Assert;
import org.junit.Test;

import java.util.Calendar;

public class ProjectTest {

    @Test
    public void test() throws Exception {
        Project project = EZExcel.toJava(ProjectTest.class.getResource("/").getFile() + "project.xlsx", Project.class);
        Assert.assertEquals(project.name, "XXX项目");
        Assert.assertEquals(project.totalPrice.toString(), "710.3");
        Assert.assertEquals(project.modules.get(3).level, 4);
        Assert.assertEquals(project.modules.get(0).process.percentage, 1.0);
        Assert.assertTrue(project.modules.get(0).process.isFinish);
        Assert.assertFalse(project.modules.get(1).process.isFinish);
        Assert.assertEquals(project.modules.get(1).process.status.size(), 2);
        Assert.assertEquals(project.modules.get(1).process.status.get(0), "是");
        Assert.assertEquals(project.modules.get(1).process.status.get(1), "否");
    }
}
