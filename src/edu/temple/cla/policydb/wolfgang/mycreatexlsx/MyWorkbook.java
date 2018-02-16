/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.temple.cla.policydb.wolfgang.mycreatexlsx;

import java.io.Closeable;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import org.apache.log4j.Logger;

/**
 * This class represents an Excel 2007 workbook.  It is limited to a
 * single worksheet.
 * @author Paul Wolfgang
 */
public class MyWorkbook implements Closeable {

    private static final Logger LOGGER = Logger.getLogger(MyWorkbook.class);

    /** The zip output stream to hold the Excel workbook */
    private ZipOutputStream zos = null;

    /** The current zip entry */
    private ZipEntry zipEntry = null;

    /**
     * Create a new MyWorkbook with no output stream
     * Used for testing only
     */
    MyWorkbook() {};


    /**
     * Create a new MyWorkbook with the given file name
     * @param fileName The name of the file to contain the Excel workbook
     * @throws java.io.FileNotFoundException
     */
    public MyWorkbook(String fileName) throws FileNotFoundException {
            this(new FileOutputStream(fileName));
            LOGGER.debug("New workbook " + fileName);
    }

    /**
     * Create a new MyWorkbook with the given OutputStream
     * @param os The OutputStream where the workbook is written.
     */
    public MyWorkbook(OutputStream os) {
        try {
            zos = new ZipOutputStream(os);
            addPartFromTemplate("[Content_Types].xml");
            addPartFromTemplate("_rels/.rels");
            addPartFromTemplate("docProps/app.xml");
            addPartFromTemplate("docProps/core.xml");
            addPartFromTemplate("xl/workbook.xml");
            addPartFromTemplate("xl/styles.xml");
            addPartFromTemplate("xl/_rels/workbook.xml.rels");
            addPartFromTemplate("xl/theme/theme1.xml");
            LOGGER.debug("initialized new workbook");
        } catch (IOException ex) {
            LOGGER.error("IOException", ex);
        }
    }

    @Override
    public void close() {
        try {
            if (zipEntry != null) {
                zos.closeEntry();
                zipEntry = null;
            }
            if (zos != null) {
                zos.close();
            }
        } catch (IOException ex) {
            LOGGER.error("Error closing stream", ex);
        }
    }

    private void addPartFromTemplate(String partName) throws IOException {
        InputStream in = findPartInTemplate(partName);
        if (in != null) {
            zipEntry = new ZipEntry(partName);
            zos.putNextEntry(zipEntry);
            copyStream(in, zos);
            zos.closeEntry();
            zipEntry = null;
        } else {
            LOGGER.error("Could not find " + partName);
        }
    }

    InputStream findPartInTemplate(String partName) throws IOException {
        Class myClass = getClass();
        String resourceName = "templates/" + partName;
        return myClass.getResourceAsStream(resourceName);
    }

    private void copyStream(InputStream in, OutputStream out) throws IOException {
        int c;
        while ((c = in.read()) != -1) {
            out.write(c);
        }
    }

    public MyWorksheet getWorksheet() {
        try {
            zipEntry = new ZipEntry("xl/worksheets/sheet1.xml");
            zos.putNextEntry(zipEntry);
            return new MyWorksheet(zos);
        } catch (IOException ex) {
            LOGGER.error("Unable to add worksheet", ex);
            return null;
        }
    }
}
