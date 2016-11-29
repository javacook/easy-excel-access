package com.javacook.easyexcellaccess;

import com.javacook.easyexcelaccess.ExcelCoordinateAccessor;
import com.javacook.easyexcelaccess.ExcelCoordinateSequencer;

import java.io.File;
import java.io.IOException;

/**
 * Created by vollmer on 13.09.16.
 */
public class ExcelCoordinateAccessorMain {

    public static void main(String[] args) throws IOException {
        ExcelCoordinateAccessor excel = new ExcelCoordinateAccessor(new File(
                "/Users/vollmer/Documents/Entwicklung/Software/javacook/my-generator/src/main/resources/MyTests.xls"));

        new ExcelCoordinateSequencer()
                .from(excel.find("Abfrageperiode").addRow(2).setCol("D"))
                .height(3).width(1)
                .sequence()
                .forEach(System.out::println);
    }

}