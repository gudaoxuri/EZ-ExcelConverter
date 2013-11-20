package com.ecfront.easybi.excelconverter.test.tojava.project;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;
import com.ecfront.easybi.excelconverter.exchange.annotation.Sheet;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

@Sheet("项目模块列表")
public class Project {

    @Map2J(value = "A1")
    public String name;

    @Map2J(value = "A4:G(-2)")
    public List<Module> modules = new ArrayList<Module>();

    @Map2J("B(-1)")
    public BigDecimal totalPrice;

}
