package com.javacook.easyexcelaccess;

import com.javacook.coordinate.CoordinateInterface;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

/**
 * A collection of routines to access Excel content very easily.
 */
public class ExcelAccessor implements ExcelEasyAccess {

    public final static double EPSILON = 1E-10;
    public final static String DEFAULT_DATE_FORMAT = "yyyy-MM-dd'T'HH:mm:ssX"; // X fuer ISO-8601
    public final static TimeZone DEFAULT_TIME_ZONE = TimeZone.getTimeZone("UTC");
    final HSSFWorkbook workbook;
    
    /**
     * Constructor to access the Ecxel file as class path resource
     * @param resourceName name of the resource
     * @throws IOException if the resource file does not exist or is not accessable
     */
    public ExcelAccessor(String resourceName) throws IOException {
        try (InputStream is = ClassLoader.getSystemResourceAsStream(resourceName)) {
            if (is == null) {
                throw new IllegalArgumentException("Resource '" + resourceName + " does not exist.");
            }
            workbook = new HSSFWorkbook(is);
        }
    }

    public ExcelAccessor(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            workbook = new HSSFWorkbook(fis);
        }
    }

    /**
     * Checks whether the excel cell at coordinate (x,y) is empty
     * @param sheetNo Excel sheet number starting with 0
     * @param x
     * @param y
     * @return true if the cell is empty
     */
    public boolean isEmpty(int sheetNo, int x, int y) {
        HSSFSheet sheet = workbook.getSheetAt(sheetNo);
        if (sheet == null) return true;
        HSSFRow row = sheet.getRow(y);
        if (row == null) return true;
        HSSFCell cell = row.getCell(x);
        if (cell == null) return true;
        return cell.getCellType() == CellType.BLANK;
    }

    public boolean isRowEmpty(int sheetNo, int row) {
        for (int col = 0; col < noCols(sheetNo); col++) {
            if (!isEmpty(sheetNo, col, row)) return false;
        }
        return true;
    }

    public boolean isOutside(int sheetNo, int x, int y) {
        return x > noCols(sheetNo) || y > noRows(sheetNo);
    }

    public Object readCell(int sheetNo, int x, int y) {
        if (isEmpty(sheetNo, x, y)) return null;
        try {
            HSSFCell cell = workbook.getSheetAt(sheetNo).getRow(y).getCell(x);
            return readCell(cell, cell.getCellType());
        } catch (Exception e) {
            throw new RuntimeException("Access to cell (sheet="+sheetNo+", x="+x+", y="+y+") failed", e);
        }
    }

    Object readCell(HSSFCell cell, CellType cellType) {
        if (cell == null) return null;
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                final double numericCellValue = cell.getNumericCellValue();
                if (DateUtil.isCellDateFormatted(cell)) {
                    return DateUtil.getJavaDate(numericCellValue, false, DEFAULT_TIME_ZONE);
                }
                return convertNumericToObject(numericCellValue);
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case BLANK:
                return null;
            case ERROR:
                throw new IllegalArgumentException("Cell is of type error.");
            case FORMULA:
                final CellType cachedFormulaResultType = cell.getCachedFormulaResultType();
                if (cachedFormulaResultType == CellType.FORMULA) {
                    throw new IllegalStateException("Cached cell type is again of type formular.");
                }
                return readCell(cell, cell.getCachedFormulaResultType());
            default:
                throw new IllegalStateException("Invalid cell type: " + cell.getCellType());
        }
    }


    public Object readCell(int sheet, CoordinateInterface coord) {
        return readCell(sheet, coord.x(), coord.y());
    }

    private Object convertNumericToObject(double  cellValue) {
        long cellValueRounded = Math.round(cellValue);
        if (Math.abs(cellValue - (double)cellValueRounded) < EPSILON) {
            if (Integer.MIN_VALUE <= cellValueRounded && cellValue <= Integer.MAX_VALUE) {
                return (int)cellValueRounded;
            }
            else {
                return cellValueRounded;
            }
        }
        else return cellValue;
    }

    @SuppressWarnings("unchecked")
    public <T> T readCell(int sheetNo, int x, int y, Class<T> clazz) {
        if (isEmpty(sheetNo, x, y)) return null;
        try {
            HSSFCell cell = workbook.getSheetAt(sheetNo).getRow(y).getCell(x);
            if (cell == null) return null;
            switch (cell.getCellType()) {
                case STRING:
                    String str = cell.getStringCellValue();
                    if (clazz == String.class) return (T) str;
                    if (clazz == Integer.class) return (T) new Integer(str);
                    if (clazz == Long.class) return (T) new Long(str);
                    if (clazz == Boolean.class) return (T) new Boolean(str);
                    if (clazz == Date.class) {
                        final SimpleDateFormat isoFormat = new SimpleDateFormat(DEFAULT_DATE_FORMAT);
                        isoFormat.setTimeZone(DEFAULT_TIME_ZONE);
                        return (T) isoFormat.parse(str);
                    }
                    throw new IllegalArgumentException("Invalid class '" + clazz + "'");
                case NUMERIC:
                    double dbl = cell.getNumericCellValue();
                    if (clazz == String.class) return (T) String.valueOf(dbl);
                    if (clazz == Integer.class) return (T) new Integer((int) dbl);
                    if (clazz == Long.class) return (T) new Long((long) dbl);
                    if (clazz == Date.class) return (T) DateUtil.getJavaDate(dbl);
                    throw new IllegalArgumentException("Invalid class '" + clazz + "'");
                case BOOLEAN:
                    boolean bool = cell.getBooleanCellValue();
                    if (clazz == String.class) return (T) String.valueOf(bool);
                    if (clazz == Integer.class) return (T) (bool ? Integer.valueOf(1) : Integer.valueOf(0));
                    if (clazz == Long.class) return (T) (bool ? Long.valueOf(1) : Long.valueOf(0));
                    throw new IllegalArgumentException("Invalid class '" + clazz + "'");
                case BLANK:
                    return null;
                case ERROR:
                    throw new IllegalArgumentException("Cell is of type error.");
                case FORMULA:
                    throw new IllegalArgumentException("Cell is of type formular.");
                default:
                    throw new IllegalStateException("Invalid cell type: " + cell.getCellType());
            }
        } catch (Exception e) {
            throw new RuntimeException("Access to cell (sheet="+sheetNo+", x="+x+", y="+y+") failed", e);
        }
    }

    public <T> T readCell(int sheet, CoordinateInterface coord, Class<T> clazz) {
        return readCell(sheet, coord.x(), coord.y(), clazz);
    }

    @Override
    public int noRows(int sheetNo) {
        return workbook.getSheetAt(sheetNo).getLastRowNum() + 1;
    }

    @Override
    public int noCols(int sheetNo) {
        int maxLastCellNum = 0;
        HSSFSheet sheet = workbook.getSheetAt(sheetNo);
        for (int i = 0; i < noRows(sheetNo); i++) {
            HSSFRow row = sheet.getRow(i);
            if (row != null) {
                short lastCellNum = sheet.getRow(i).getLastCellNum();
                if (lastCellNum > maxLastCellNum) maxLastCellNum = lastCellNum;
            }
        }
        return maxLastCellNum;
    }

    @Override
    public String sheetName(int sheetNo) {
        HSSFSheet sheet = workbook.getSheetAt(sheetNo);
        return sheet.getSheetName();
    }

}
