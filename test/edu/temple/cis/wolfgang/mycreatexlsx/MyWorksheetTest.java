/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.temple.cis.wolfgang.mycreatexlsx;

import java.util.TimeZone;
import java.util.GregorianCalendar;
import java.util.Date;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author Paul Wolfgang
 */
public class MyWorksheetTest {

    public MyWorksheetTest() {
    }

    @BeforeClass
    public static void setUpClass() throws Exception {
    }

    @AfterClass
    public static void tearDownClass() throws Exception {
    }

    @Before
    public void setUp() {
    }

    @After
    public void tearDown() {
    }

    /**
     * Test of toBase26 method, of class MyWorksheet.
     */
    @Test
    public void testToBase26() {
        System.out.println("toBase26");
        assertEquals("A", MyWorksheet.toBase26(0, ""));
        assertEquals("Z", MyWorksheet.toBase26(25, ""));
        assertEquals("AA", MyWorksheet.toBase26(26, ""));
        assertEquals("AB", MyWorksheet.toBase26(27, ""));
        assertEquals("BA", MyWorksheet.toBase26(52, ""));
        assertEquals("ZZ", MyWorksheet.toBase26(701, ""));
        assertEquals("AAA", MyWorksheet.toBase26(702, ""));
        assertEquals("AAB", MyWorksheet.toBase26(703, ""));
    }


    @Test public void testExcelDate1() {
        GregorianCalendar c = new GregorianCalendar();
        c.setTimeZone(TimeZone.getTimeZone("GMT"));
        c.set(1979, 3, 2);
        Date d = c.getTime();
        assertEquals("1979-04-02", String.format("%tF", d));
        assertEquals("1979-04-02", 28947, MyWorksheet.computeExcelDate(d));
    }


    @Test public void testExcelDate2() {
        GregorianCalendar c = new GregorianCalendar();
        c.setTimeZone(TimeZone.getTimeZone("GMT"));
        c.set(1979, 3, 3);
        Date d = c.getTime();
        assertEquals("1979-04-03", String.format("%tF", d));
        assertEquals("1979-04-03", 28948, MyWorksheet.computeExcelDate(d));
    }

}