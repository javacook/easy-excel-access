package com.javacook.easyexcelaccess;

import com.javacook.coordinate.Coordinate;
import com.javacook.coordinate.CoordinateInterface;

/**
 * Represents an Excel coordinate by a column and row number (1,...,n). The
 * column can also be defined by a letter combination (A,...,Z,AA,...)
 */
public class ExcelCoordinate implements CoordinateInterface {

    public static final int COL_MAX = 65536;
    public static final int ROW_MAX = 65536;

    protected CoordinateInterface c;

    /**
     * Standard constructor
     * @param col numeric column coordinate starting with 1
     * @param row numeric row coordinate starting with 1
     */
    public ExcelCoordinate(int col, int row) {
        this(new Coordinate(col-1, row-1));
    }

    /**
     * Standard constructor using a letter combination
     * @param col letter combination starting with 'A'
     * @param row numeric row coordinate starting with 1
     */
    public ExcelCoordinate(String col, int row) {
        this(ExcelUtil.calculateColNo(col),row);
    }

    public ExcelCoordinate(CoordinateInterface coord) {
        c = coord;
    }

    public int col() {
        return x() + 1;
    }

    public int row() {
        return y() + 1;
    }

    public ExcelCoordinate setCol(int newCol) {
        return new ExcelCoordinate(newCol, row());
    }

    public ExcelCoordinate setCol(String newCol) {
        return new ExcelCoordinate(ExcelUtil.calculateColNo(newCol), row());
    }

    public ExcelCoordinate setRow(int newRow) {
        return new ExcelCoordinate(col(), newRow);
    }

    public ExcelCoordinate setRow(String newRow) {
        return new ExcelCoordinate(col(), ExcelUtil.calculateColNo(newRow));
    }

    public ExcelCoordinate addRow(int diff) {
        return new ExcelCoordinate(col(), row() + diff);
    }

    public ExcelCoordinate add(int colDiff, int rowDiff) {
        return new ExcelCoordinate(col() + colDiff, row() + rowDiff);
    }

    public ExcelCoordinate subRow(int diff) {
        return new ExcelCoordinate(x(), y() - diff);
    }

    public ExcelCoordinate addCol(int diff) {
        return new ExcelCoordinate(col() + diff, row());
    }

    public ExcelCoordinate subCol(int diff) {
        return new ExcelCoordinate(col() - diff, row());
    }

    public ExcelCoordinate incRow() {
        return addRow(1);
    }

    public ExcelCoordinate incCol() {
        return addCol(1);
    }

    public ExcelCoordinate decRow() {
        return subRow(1);
    }

    public ExcelCoordinate decCol() {
        return subCol(1);
    }

    public boolean isLeftFrom(CoordinateInterface coord) {
        return x() < coord.x();
    }

    public boolean isRightFrom(CoordinateInterface coord) {
        return x() > coord.x();
    }

    public boolean isOver(CoordinateInterface coord) {
        return y() < coord.y();
    }

    public boolean isUnder(CoordinateInterface coord) {
        return y() > coord.y();
    }


    public boolean hasSameRow(CoordinateInterface coord) {
        return y() == coord.y();
    }

    public boolean hasSameCol(CoordinateInterface coord) {
        return x() == coord.x();
    }

    @Override
    public String toString() {
        int col = col();
        String letters = ExcelUtil.calculateColLetters(col);
        return "[col=" + letters + " ,row="+ row() + "]";
    }


    @Override
    public int x() {
        return c.x();
    }

    @Override
    public int y() {
        return c.y();
    }
}
