package com.javacook.easyexcellaccess;

/**
 * Created by vollmer on 30.07.16.
 */
public interface ExcelEasyAccess {

    <T> T readCell(int sheet, int x, int y, Class<T> clazz);

}