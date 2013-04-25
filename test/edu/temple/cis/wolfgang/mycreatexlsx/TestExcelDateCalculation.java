/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package edu.temple.cis.wolfgang.mycreatexlsx;

import java.util.Arrays;
import java.util.ArrayList;
import java.util.List;
import java.util.TimeZone;
import java.util.GregorianCalendar;
import java.util.Collection;
import java.util.Date;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;
import static org.junit.Assert.assertEquals;

/**
 *
 * @author Paul Wolfgang
 */
@RunWith(value = Parameterized.class)
public class TestExcelDateCalculation {

    private String dateString;
    private long expected;


    @Parameters
    public static Collection<Object[]> getParameters() {
        List<Object[]> params = new ArrayList<Object[]>(
                Arrays.asList(new Object[][]{
                    {"1900-01-01", 1},
                    {"1900-01-02", 2},
                    {"1900-01-31", 31},
                    {"1900-02-01", 32},
                    {"1900-02-28", 59},
                    {"1900-03-01", 61},
                    {"1900-07-10", 192},
                    {"1901-01-01", 367},
                    {"1901-01-02", 368},
                    {"1901-01-31", 397},
                    {"1901-02-01", 398},
                    {"1901-02-28", 425},
                    {"1901-03-01", 426},
                    {"1901-07-10", 557},
                    {"1979-01-01", 28856},
                    {"1979-01-02", 28857},
                    {"1979-01-31", 28886},
                    {"1979-02-01", 28887},
                    {"1979-02-28", 28914},
                    {"1979-03-01", 28915},
                    {"1979-03-02", 28916},
                    {"1979-07-09", 29045},
                    {"1979-07-10", 29046}
        }));
        GregorianCalendar calendar1979_03_01 = new GregorianCalendar();
        calendar1979_03_01.setTimeZone(TimeZone.getTimeZone("GMT"));
        calendar1979_03_01.clear();
        calendar1979_03_01.set(1979, 2, 1);
        Date date1979_03_01 = calendar1979_03_01.getTime();
        long goodDateFor1979_03_01 = 28915;
        long msPerDay = 24 * 3600 * 1000;
        for (int i = 0; i < 150; i++) {
            Date d = new Date(date1979_03_01.getTime() + i * msPerDay);
            String dateString = String.format("%tF", d.getTime()+ 5*3600*1000);
            Long expected = goodDateFor1979_03_01 + i;
            params.add(new Object[]{dateString, expected});
        }
        return params;
    }

    public TestExcelDateCalculation(String dateString, long expected) {
        this.dateString = dateString;
        this.expected = expected;
    }

    @Test
    public void testDateCalculation() throws Exception {
        long actual = MyWorksheet.computeExcelDate(dateString);
        assertEquals(dateString, expected, actual);
    }

}
