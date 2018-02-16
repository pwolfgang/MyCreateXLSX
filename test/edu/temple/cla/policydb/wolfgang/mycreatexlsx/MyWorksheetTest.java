/* 
 * Copyright (c) 2018, Temple University
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * * Redistributions of source code must retain the above copyright notice, this
 *   list of conditions and the following disclaimer.
 * * Redistributions in binary form must reproduce the above copyright notice,
 *   this list of conditions and the following disclaimer in the documentation
 *   and/or other materials provided with the distribution.
 * * All advertising materials features or use of this software must display 
 *   the following  acknowledgement
 *   This product includes software developed by Temple University
 * * Neither the name of the copyright holder nor the names of its 
 *   contributors may be used to endorse or promote products derived 
 *   from this software without specific prior written permission. 
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 */
package edu.temple.cla.policydb.wolfgang.mycreatexlsx;

import edu.temple.cla.policydb.wolfgang.mycreatexlsx.MyWorksheet;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;
import static edu.temple.cla.policydb.wolfgang.mycreatexlsx.MyWorksheet.computeExcelDate;

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
        assertEquals("A", MyWorksheet.toBase26(0));
        assertEquals("Z", MyWorksheet.toBase26(25));
        assertEquals("AA", MyWorksheet.toBase26(26));
        assertEquals("AB", MyWorksheet.toBase26(27));
        assertEquals("BA", MyWorksheet.toBase26(52));
        assertEquals("ZZ", MyWorksheet.toBase26(701));
        assertEquals("AAA", MyWorksheet.toBase26(702));
        assertEquals("AAB", MyWorksheet.toBase26(703));
    }


    
    @Test public void testExcelDate3() {
        assertEquals(1, computeExcelDate("1900-01-01"));
        assertEquals(58, computeExcelDate("1900-02-27"));
        assertEquals(59, computeExcelDate("1900-02-28"));
        assertEquals(61, computeExcelDate("1900-03-01"));
    }

}