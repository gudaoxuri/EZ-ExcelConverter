package com.ecfront.easybi.excelconverter.test.tojava.project;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;

import java.util.Date;

public class Process {

    @Map2J("E")
    private Double percentage;
    @Map2J(value = "F", trueValue = "是")
    private Boolean isFinish;
    @Map2J("G")
    private Date finishDate;

}
