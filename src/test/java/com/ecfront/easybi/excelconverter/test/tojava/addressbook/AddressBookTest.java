package com.ecfront.easybi.excelconverter.test.tojava.addressbook;


import com.ecfront.easybi.excelconverter.exchange.EZExcel;
import junit.framework.Assert;
import org.junit.Test;

public class AddressBookTest {
    @Test
    public void test() throws Exception {
        AddressBook addressBook = EZExcel.toJava(AddressBookTest.class.getResource("/").getFile() + "address book.xlsx", AddressBook.class);
        Assert.assertEquals(addressBook.companyName, "XXX公司通讯录");
        Assert.assertEquals(addressBook.items.get(1).name, "李四");
    }
}
