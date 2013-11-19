package com.ecfront.easybi.excelconverter.test.tojava.project;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;

import java.math.BigDecimal;

public class Module {

    @Map2J("A")
    public String module;
    @Map2J("B")
    public String function;
    @Map2J("C")
    public int level;
    @Map2J("D")
    public BigDecimal price;
    @Map2J(value = "E:G")
    public Process process = new Process();

}
