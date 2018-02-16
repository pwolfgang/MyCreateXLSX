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

import edu.temple.cla.policydb.wolfgang.mycreatexlsx.MyWorkbook;
import java.io.InputStream;
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
public class MyWorkbookTest {

    public MyWorkbookTest() {
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

    @Test
    public void testFindPartInTemplate() throws Exception {
        MyWorkbook instance = new MyWorkbook();
        InputStream in = instance.findPartInTemplate("[Content_Types].xml");
        assertNotNull("Could not find \"[Content_Types].xml\"",in);
        in = instance.findPartInTemplate("_rels/.rels");
        assertNotNull("Could not find \"_rels/.rels\"", in);
        in = instance.findPartInTemplate("docProps/app.xml");
        assertNotNull("Could not findi\"docProps/app.xml\"", in);
        in = instance.findPartInTemplate("docProps/core.xml");
        assertNotNull("Could not find \"docProps/core.xml\"", in);
        in = instance.findPartInTemplate("xl/workbook.xml");
        assertNotNull("Could not find \"xl/workbook.xml\"", in);
        in = instance.findPartInTemplate("xl/styles.xml");
        assertNotNull("Could not find \"xl/styles.xml\"", in);
        in = instance.findPartInTemplate("xl/_rels/workbook.xml.rels");
        assertNotNull("Could not find \"xl/_rels/workbook.xml.rels\"", in);
        in = instance.findPartInTemplate("xl/theme/theme1.xml");
        assertNotNull("Could not find \"xl/theme/theme1.xml\"", in);

    }
}