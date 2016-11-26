package com.javacook.easyexcelaccess;

/**
 * Created by vollmer on 07.09.16.
 */
public class ExcelUtil {

    public static int calculateColNo(final String excelCol) {
        int result = 0;
        for (char c : excelCol.toUpperCase().toCharArray()) {
            result = 26 * result;
            result += (int)c - (int)'A' + 1;
        }
        return result;
    }

    public static String calculateColLetters(final int excelCol) {
        String letters = "";
        int rest = excelCol - 1;
        do {
            letters = (char)(rest % 26 + 'A') + letters;
            rest = rest / 26 - 1;
        } while (rest >= 0);

        return letters;
    }

}
