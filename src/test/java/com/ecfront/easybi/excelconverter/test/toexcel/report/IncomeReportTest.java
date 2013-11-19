package com.ecfront.easybi.excelconverter.test.toexcel.report;

import com.ecfront.easybi.excelconverter.exchange.EZExcel;
import org.junit.Test;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

public class IncomeReportTest {
    @Test
    public void test() throws Exception {
        IncomeReport report = new IncomeReport();
        report.head = new HashMap() {{
            put("xxx收支汇总", new ArrayList() {{
                add("年份");
                add(new HashMap() {{
                    put("收入", new Object[]{
                            "工资",
                            new HashMap() {{
                                put("福利", new String[]{"餐补", "通讯补贴"});
                            }}
                    });
                }});
                add("支出");
            }});
        }};
        report.side = new ArrayList<Integer>() {{
            add(201306);
            add(null);
            add(201308);
        }};
        report.data = new ArrayList() {{
            add(new Double[]{10000D, 100D, 44D, 300D});
            add(new Double[]{10000D, 200D, 300D});
            add(new Double[]{null, 200D, 300D});
            add(new BigDecimal[]{new BigDecimal(20000D), new BigDecimal(20200D), new BigDecimal(21000D)});
        }};
        report.fill = new String[]{"a", "b"};
        //EZExcel.toExcel(report, "D:\\workbook" + new Date().getTime() + ".xlsx");
        EZExcel.toExcel(report, "D:\\workbook" + new Date().getTime() + ".xlsx", IncomeReportTest.class.getResource("/").getFile() + "template.xlsx");
    }
}
