package com.ecfront.easybi.excelconverter.test.tojava.addressbook;


import com.ecfront.easybi.excelconverter.exchange.annotation.Map2J;
import com.ecfront.easybi.excelconverter.exchange.annotation.Sheet;

import java.util.ArrayList;
import java.util.List;

@Sheet
public class AddressBook {
    @Map2J("A1")
    public String companyName;
    @Map2J("A3:G(-1)")
    public List<Item> items = new ArrayList<Item>();
}
