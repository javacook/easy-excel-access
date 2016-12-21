package com.javacook.easyexcelaccess;

import com.javacook.coordinate.CoordinateInterface;

import java.io.File;
import java.io.IOException;

/**
 * Created by vollmer on 28.08.16.
 */
public class ExcelCoordinateAccessor extends ExcelAccessor {

    public final int defaultSheet;

    public ExcelCoordinateAccessor(String resourceName) throws IOException {
        this(resourceName, 0);
    }

    public ExcelCoordinateAccessor(File file) throws IOException {
        this(file, 0);
    }

    /**
     * @param resourceName
     * @param defaultSheet sheet number starting with 0
     * @throws IOException
     */
    public ExcelCoordinateAccessor(String resourceName, int defaultSheet) throws IOException {
        super(resourceName);
        this.defaultSheet = defaultSheet;
    }

    /**
     *
     * @param file
     * @param defaultSheet sheet number starting with 0
     * @throws IOException
     */
    public ExcelCoordinateAccessor(File file, int defaultSheet) throws IOException {
        super(file);
        this.defaultSheet = defaultSheet;
    }


    public boolean isEmpty(int sheetNo, CoordinateInterface coordinate) {
        return isEmpty(sheetNo, coordinate.x(), coordinate.y());
    }

    public boolean isOutside(int sheetNo, CoordinateInterface coordinate) {
        return isOutside(sheetNo, coordinate.x(), coordinate.y());
    }

    public Object read(int sheetNo, CoordinateInterface coord) {
        return readCell(sheetNo, coord.x(), coord.y());
    }

    public String readString(int sheetNo, CoordinateInterface coord) {
        return readCell(sheetNo, coord.x(), coord.y(), String.class);
    }

    public Integer readInteger(int sheetNo, CoordinateInterface coord) {
        return readCell(sheetNo, coord.x(), coord.y(), Integer.class);
    }

    public Long readLong(int sheetNo, CoordinateInterface coord) {
        return readCell(sheetNo, coord.x(), coord.y(), Long.class);
    }

    public Boolean readBoolean(int sheetNo, CoordinateInterface coord) {
        return readCell(sheetNo, coord.x(), coord.y(), Boolean.class);
    }


    public ExcelCoordinate find(int sheetNo, String word) {
        return find(sheetNo, word, false);
    }

    public ExcelCoordinate findInColumn(int sheetNo, int col, String word, boolean nullInsteadOfException) {
            for (int row = 1; row <= noRows(sheetNo); row++) {
                if (contains(word, sheetNo, col, row)) return new ExcelCoordinate(col, row);
            }
        if (nullInsteadOfException) return null;
        throw new RuntimeException("Could not find '" +  word +  "' in sheet " + sheetNo +
                " and column " + col);
    }

    public ExcelCoordinate find(int sheetNo, String word, boolean nullInsteadOfException) {
        for (int col = 1; col <= noCols(sheetNo); col++) {
            for (int row = 1; row <= noRows(sheetNo); row++) {
                if (contains(word, sheetNo, col, row)) return new ExcelCoordinate(col, row);
            }
        }
        if (nullInsteadOfException) return null;
        throw new RuntimeException("Could not find '" +  word +  "' in sheet " + sheetNo);
    }

    protected boolean contains(String word, int sheetNo, int col, int row) {
        String content = readCell(sheetNo, col-1, row-1, String.class);
        return ((content == null && word == null) || (content != null && content.contains(word)));
    }


    /*---------------------------*\
     * For default sheet         *
    \*---------------------------*/

    public boolean isEmpty(CoordinateInterface coord) {
        return isEmpty(defaultSheet, coord);
    }

    public boolean isOutside(CoordinateInterface coord) {
        return isOutside(defaultSheet, coord);
    }

    public Object read(CoordinateInterface coord) {
        return read(defaultSheet, coord);
    }

    public String readString(CoordinateInterface coord) {
        return readString(defaultSheet, coord);
    }

    public Integer readInteger(CoordinateInterface coord) {
        return readInteger(defaultSheet, coord);
    }

    public Long readLong(CoordinateInterface coord) {
        return readLong(defaultSheet, coord);
    }

    public Boolean readBoolean(CoordinateInterface coord) {
        return readBoolean(defaultSheet, coord);
    }

    public ExcelCoordinate find(String word) {
        return find(defaultSheet, word);
    }

    public ExcelCoordinate find(String word, boolean nullInsteadOfException) {
        return find(defaultSheet, word, nullInsteadOfException);
    }


}
