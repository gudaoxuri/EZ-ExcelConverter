package com.ecfront.easybi.excelconverter.test.tojava.project;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;

import java.util.Date;
import java.util.List;

public class Process {

    @Map2J("E")
    public Double percentage;
    @Map2J(value = "F", trueValue = "是")
    public Boolean isFinish;
    @Map2J(value = "F", getStruct = true)
    public List<String> status;
    @Map2J("G")
    public Date finishDate;

}
