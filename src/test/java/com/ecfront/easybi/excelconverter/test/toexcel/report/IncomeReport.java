package com.ecfront.easybi.excelconverter.test.toexcel.report;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2E;
import com.ecfront.easybi.excelconverter.exchange.annotation.Mode;
import com.ecfront.easybi.excelconverter.exchange.annotation.Sheet;
import com.ecfront.easybi.excelconverter.exchange.annotation.Validation;

import java.util.List;
import java.util.Map;

@Sheet("收支报表")
public class IncomeReport {

    @Map2E(value = "head", abs = "B1")
    public Map head;
    @Map2E(value = "side", abs = "B6", mode = Mode.COLUMN)
    public List<Integer> side;
    @Map2E(value = "data", top = "head", left = "side", isMatrix = true)
    public List<Double[]> data;
    @Validation(top = "head", left = "data", length = "data", mode = Mode.COLUMN)
    public String[] fill;
}
