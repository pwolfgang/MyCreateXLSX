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
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.TimeZone;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.log4j.Logger;


/**
 *
 * @author Paul Wolfgang
 */
public class MyWorksheet implements Closeable {

    private static final Logger LOGGER = Logger.getLogger(MyWorkbook.class);


    private static final long BASE_TIME;
    private static final long FEB28_1900;
    private static final long MS_PER_DAY = 24 * 3600 * 1000;

    static {
        GregorianCalendar c = new GregorianCalendar();
        c.setTimeZone(TimeZone.getTimeZone("GMT"));
        c.set(1899, 11, 30, 0, 0, 0);
        BASE_TIME = c.getTime().getTime();
        c.set(1900, 1, 28, 0, 0, 0);
        FEB28_1900 = c.getTimeInMillis();
    }

    private PrintWriter w = null;
    int rowNum = 0;
    int colNum = 0;

    MyWorksheet(OutputStream os) {
        try {
            w = new PrintWriter(new OutputStreamWriter(os, "UTF-8"));
        }
        catch (UnsupportedEncodingException ex) {
            w = new PrintWriter(new OutputStreamWriter(os));
        }
        w.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        w.println("<worksheet");
        w.println("        xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
        w.println("        xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        w.println("<sheetData>");
    }

    @Override
    public void close() {
        w.println("</sheetData>");
        w.println("</worksheet>");
        w.flush();
    }

    public static long computeExcelDate(String dateString) {
        String[] dateTokens = dateString.split("-|\\s+");
        int year = Integer.parseInt(dateTokens[0]);
        int month = Integer.parseInt(dateTokens[1]);
        int day = Integer.parseInt(dateTokens[2]);
        GregorianCalendar c = new GregorianCalendar();
        c.setTimeZone(TimeZone.getTimeZone("GMT"));
        c.clear();
        c.set(year, month-1, day);
        Date d = c.getTime();
        return computeExcelDate(d);
    }

    public static long computeExcelDate(Date d) {
        long time = d.getTime();
        if (time > FEB28_1900) {
            time = time + MS_PER_DAY;
        }
        long deltaMilliseconds = time - BASE_TIME;
        long numDays = deltaMilliseconds / MS_PER_DAY;
        return numDays;
    }
    
    public void startRow() {
        rowNum++;
        colNum = 0;
        w.printf("<row r=\"%d\">%n", rowNum);
    }
    
    public void endRow() {
        w.println("</row>");
    }

    private void startColumn() {
        w.println("<c r=\"" + toBase26(colNum++) + (rowNum) + "\">");
    }

    private void endColumn() {
        w.println("</c>");
    }

    public void addCell(int x) {
        startColumn();
        w.println("<v>" + x + "</v>");
        endColumn();
    }

    public void addCell(double x) {
        startColumn();
        w.println("<v>" + x + "</v>");
        endColumn();
    }

    public void addCell(String s) {
        w.println("<c r=\"" + toBase26(colNum++) + (rowNum) + "\" t=\"inlineStr\">");
        w.println("<is><t>" + StringEscapeUtils.escapeXml11(s) + "</t></is>");
        endColumn();
    }

    public void addCell(Date d) {
        long numDays = computeExcelDate(d);
        w.println("<c r=\"" + toBase26(colNum++) + (rowNum) + "\" s=\"1\">");
        w.println("<v>" + numDays + "</v>");
        endColumn();
    }

    public void addDateCell(String s) {
        long numDays = computeExcelDate(s);
        w.println("<c r=\"" + toBase26(colNum++) + (rowNum) + "\" s=\"1\">");
        w.println("<v>" + numDays + "</v>");
        endColumn();
    }


    public static String toBase26(int x) {
        return toBase26(x, new StringBuilder()).toString();
    }

    public static StringBuilder toBase26(int x, StringBuilder prefix) {
        if (x < 26) {
            int charValue = 'A';
            int newCharValue = charValue + x;
            char newChar = (char) newCharValue;
            return prefix.append(newChar);
        } else {
            int y = x / 26 - 1;
            int z = x % 26;
            return prefix.append(toBase26(z, toBase26(y, new StringBuilder())));
        }
    }
}
