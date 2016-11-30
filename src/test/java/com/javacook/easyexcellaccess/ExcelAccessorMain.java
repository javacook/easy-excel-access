package com.javacook.easyexcellaccess;

import com.javacook.easyexcelaccess.ExcelAccessor;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

/**
 * Created by vollmer on 28.08.16.
 */
public class ExcelAccessorMain {

    public static void main(String[] args) throws IOException {
        ExcelAccessor excelAccessor = new ExcelAccessor("zwei.xls");
        System.out.println(excelAccessor.readCell(0, 0, 2, String.class));
        System.out.println(excelAccessor.readCell(0, 0, 2));
        System.out.println(excelAccessor.readCell(0, 1, 2));
        System.out.println(excelAccessor.readCell(0, 1, 0, Integer.class));
        System.out.println(excelAccessor.noRows(0));
        System.out.println(excelAccessor.noCols(0));
        final Date date = (Date)excelAccessor.readCell(0, 0, 2);
        final SimpleDateFormat isoFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssX");
        isoFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
        System.out.println(isoFormat.format(new Date()));
        System.out.println(isoFormat.format(date));
    }

}