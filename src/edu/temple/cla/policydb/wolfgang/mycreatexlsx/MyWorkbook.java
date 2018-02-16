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
