package com.javacook.easyexcellaccess;

import java.io.IOException;

import static org.junit.Assert.*;

/**
 * Created by vollmer on 28.08.16.
 */
public class ExcelAccessorMain {

    public static void main(String[] args) throws IOException {
        ExcelAccessor excelAccessor = new ExcelAccessor("/fibotest.xls");
        System.out.println(excelAccessor.readCell(0, 0, 2, String.class));
        System.out.println(excelAccessor.readCell(0, 1, 0, Integer.class));
    }

}