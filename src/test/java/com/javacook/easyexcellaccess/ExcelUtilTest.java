package com.javacook.easyexcellaccess;

import com.javacook.easyexcelaccess.ExcelUtil;
import org.junit.Test;

import static org.junit.Assert.assertEquals;

/**
 * Created by vollmer on 09.10.16.
 */
public class ExcelUtilTest {
    @Test
    public void calculateColNo() throws Exception {
        assertEquals(1, ExcelUtil.calculateColNo("A"));
        assertEquals(2, ExcelUtil.calculateColNo("B"));
        assertEquals(26, ExcelUtil.calculateColNo("Z"));
        assertEquals(27, ExcelUtil.calculateColNo("AA"));
        assertEquals(52, ExcelUtil.calculateColNo("AZ"));
        assertEquals(53, ExcelUtil.calculateColNo("BA"));
    }

    @Test
    public void calculateColLetters() throws Exception {
        assertEquals("A", ExcelUtil.calculateColLetters(1));
        assertEquals("B", ExcelUtil.calculateColLetters(2));
        assertEquals("Z", ExcelUtil.calculateColLetters(26));
        assertEquals("AA", ExcelUtil.calculateColLetters(27));
        assertEquals("AZ", ExcelUtil.calculateColLetters(52));
        assertEquals("BA", ExcelUtil.calculateColLetters(53));
    }

}