package com.javacook.easyexcelaccess;

/**
 * Base routines to access Excel content.
 */
public interface ExcelEasyAccess {

    /**
     * Reads the content of the cell at the coordintes <code>x</code> and <code>y</code>
     * from the sheet with no <code>sheet</code>. The type of the Excel cell content is 
     * automatically detected and the class of the result is respecting String, Number, 
     * Boolean or Date.
     * @param sheet Excel sheet number starting with 0
     * @param x column coordinate starting with 0
     * @param y row coordinate starting with 0
     * @return cell content
     */
    Object readCell(int sheet, int x, int y);

    /**
     * Reads the content of the cell at the coordintes <code>x</code> and <code>y</code>
     * from the sheet with no <code>sheet</code>.
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
     * @return Row size of bounding box containing real data
     */
    int noRows(int sheet);

    /**
     * Maximal row column number containing real data plus 1
     * @param sheet Excel sheet number starting with 0
     * @return Column size of bounding box containing real data
     */
    int noCols(int sheet);

    /**
     * Yields the name of the sheet with index <code>sheet</code>
     * @param sheet index of the sheet.
     * @return name of the sheet with index <code>sheet</code>
     */
    String sheetName(int sheet);

}