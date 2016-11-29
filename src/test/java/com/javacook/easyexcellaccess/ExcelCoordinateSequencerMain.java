package com.javacook.easyexcellaccess;

import com.javacook.easyexcelaccess.ExcelCoordinate;
import com.javacook.easyexcelaccess.ExcelCoordinateSequencer;
import com.javacook.easyexcelaccess.ExcelUtil;

/**
 * Created by vollmer on 13.09.16.
 */
public class ExcelCoordinateSequencerMain {

    public static void main(String[] args) {

        System.out.println(ExcelUtil.calculateColNo("F"));

        new ExcelCoordinateSequencer()
                .fromCol(3).fromRow(5).toCol(5).toRow(6)
                .sequence()
                .forEach(System.out::println);

        new ExcelCoordinateSequencer()
                .fromCol(3).fromRow(5).toCol(5).toRow(6)
                .forEach((coord, i) -> System.out.println(coord.col() + " - " + i));

        new ExcelCoordinateSequencer()
                .from(2, 3).to(7, 9)
                .stopWhen(coord -> coord.col() + coord.row() > 10)
                .forEach(System.out::println);


        new ExcelCoordinateSequencer()
                .from(2, 3).to(7, 9)
                .stopWhen(coord -> coord.hasSameCol(new ExcelCoordinate(4,8)))
                .forEach(System.out::println);

        new ExcelCoordinateSequencer()
                .from(2, 3).to(7, 9)
                .stopWhen(coord -> coord.hasSameRow(new ExcelCoordinate(4,8)))
                .forEach(System.out::println);
    }

}