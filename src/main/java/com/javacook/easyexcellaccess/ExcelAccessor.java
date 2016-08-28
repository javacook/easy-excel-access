package com.javacook.easyexcellaccess;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;

public class ExcelAccessor implements ExcelEasyAccess {

    final HSSFWorkbook workbook;

    public ExcelAccessor(String resourceName) throws IOException {
        InputStream is = Objects.requireNonNull(getClass().getResourceAsStream(resourceName),
                "Resource '" + resourceName + " does not exist.");
        workbook = new HSSFWorkbook(is);
    }

    public ExcelAccessor(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        workbook = new HSSFWorkbook(fis);
    }

    public boolean isEmpty(int sheetNo, int x, int y) {
        HSSFSheet sheet = workbook.getSheetAt(sheetNo);
        return sheet.getRow(y).getCell(x) == null;
    }

    public <T> T readCell(int sheetNo, int x, int y, Class<T> clazz) {
        HSSFSheet sheet = workbook.getSheetAt(sheetNo);
        HSSFRow row = sheet.getRow(y);
        HSSFCell cell = row.getCell(x);
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                String str = cell.getStringCellValue();
                if (clazz == String.class) return (T) str;
                if (clazz == Integer.class) return (T) new Integer(str);
                if (clazz == Long.class) return (T) new Long(str);
                if (clazz == Boolean.class) return (T) new Boolean(str);
                throw new IllegalArgumentException("Invalid class '" + clazz + "'");
            case Cell.CELL_TYPE_NUMERIC:
                double dbl = cell.getNumericCellValue();
                if (clazz == String.class) return (T) String.valueOf(dbl);
                if (clazz == Integer.class) return (T) new Integer((int) dbl);
                if (clazz == Long.class) return (T) new Long((long) dbl);
                throw new IllegalArgumentException("Invalid class '" + clazz + "'");
            case Cell.CELL_TYPE_BOOLEAN:
                boolean bool = cell.getBooleanCellValue();
                if (clazz == String.class) return (T) String.valueOf(bool);
                if (clazz == Integer.class) return (T) (bool ? Integer.valueOf(1) : Integer.valueOf(0));
                if (clazz == Long.class) return (T) (bool ? Long.valueOf(1) : Long.valueOf(0));
                throw new IllegalArgumentException("Invalid class '" + clazz + "'");
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_ERROR:
                throw new IllegalArgumentException("Cell is of type error.");
            case Cell.CELL_TYPE_FORMULA:
                throw new IllegalArgumentException("Cell is of type formular.");
            default:
                throw new IllegalStateException("Invalid cell type: " + cell.getCellType());
        }
    }

}
