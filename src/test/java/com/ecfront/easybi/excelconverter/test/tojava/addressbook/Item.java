package com.ecfront.easybi.excelconverter.test.tojava.addressbook;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;

public class Item {
    @Map2J("A")
    public String name;
    @Map2J("B")
    public String dept;
    @Map2J("C")
    public String phone;
}
