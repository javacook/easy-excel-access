package com.javacook.easyexcelaccess;

/**
 * Base routines to access Excel content.
 */
public interface ExcelEasyAccess {

    /**
     * Reads the content of the cell at the coordintes <code>x</code> and <code>y</code>
     * in the sheet with no <code>sheet</code>.
     *
     * @param sheet Excel sheet number starting with 0
     * @param x column coordinate starting with 0
     * @param y row coordinate starting with 0
     * @param clazz Class to convert the content to
     * @param <T>
     * @return cell content
     */
    <T> T readCell(int sheet, int x, int y, Class<T> clazz);

    /**
     * Maximal row number containing real data plus 1
     * @param sheet Excel sheet number starting with 0
     */
    int noRows(int sheet);

    /**
     * Maximal row column number containing real data plus 1
     * @param sheet Excel sheet number starting with 0
     */
    int noCols(int sheet);
}