package com.javacook.easyexcelaccess;

/**
 * Column text to number and reverse calculator
 */
public class ExcelUtil {

    /**
     * Excel text label of the column ("A"..."Z", "AA"...)
     * @param excelCol
     * @return Corresponding column number starting with 1 (= "A")
     */
    public static int calculateColNo(final String excelCol) {
        int result = 0;
        for (char c : excelCol.toUpperCase().toCharArray()) {
            result = 26 * result;
            result += (int)c - (int)'A' + 1;
        }
        return result;
    }


    /**
     * Corresponding column number starting with 1 (= "A")
     * @param excelCol
     * @return Excel text label of the column ("A"..."Z", "AA"...)
     */
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
