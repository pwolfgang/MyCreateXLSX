/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
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