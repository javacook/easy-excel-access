package com.javacook.easyexcelaccess;

import org.junit.Assert;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.*;

/**
 * Created by vollmerj on 26.01.17.
 */
public class ExcelAccessorTest {

    @Test
    public void readCell() throws Exception {
        ExcelCoordinateAccessor excel = new ExcelCoordinateAccessor("formular.xls");
        Assert.assertEquals(1, excel.readCell(0, 0, 0));
        Assert.assertEquals(2, excel.readCell(0, 0, 1));
        Assert.assertEquals(3, excel.readCell(0, 1, 0));
    }

}