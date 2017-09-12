/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package edu.temple.cis.wolfgang.mycreatexlsx;

import java.io.OutputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
import javax.sql.DataSource;
import org.apache.log4j.Logger;

/**
 *
 * @author Paul
 */
public class Util {

    private static final Logger LOGGER = Logger.getLogger(Util.class);

    /**
     * Method to add a long value to a worksheet
     * @param sheet The worksheet
     * @param i The column index
     * @param rs Result set containing the data
     * @throws SQLException
     */
    private static void addLongValue(MyWorksheet sheet, int i, ResultSet rs) throws SQLException {
        long longValue = rs.getLong(i + 1);
        if (!rs.wasNull()) {
            sheet.addCell(longValue);
        } else {
            sheet.addCell("null");
        }
    }

    /**
     * Method to add an int value to a worksheet
     * @param sheet The worksheet
     * @param i The column index
     * @param rs Result set containing the data
     * @throws SQLException
     */
    private static void addIntValue(MyWorksheet sheet, int i, ResultSet rs) throws SQLException {
        int intValue = rs.getInt(i + 1);
        if (!rs.wasNull()) {
            sheet.addCell(intValue);
        } else {
            sheet.addCell("null");
        }
    }

    public static void addColumn(int columnType, int i, MyWorksheet sheet, ResultSet rs) {
        try {
            switch (columnType) {
                case Types.BIT:
                case Types.TINYINT:
                case Types.SMALLINT:
                case Types.INTEGER:
                    addIntValue(sheet, i, rs);
                    break;
                case Types.BIGINT:
                    addLongValue(sheet, i, rs);
                    break;
                case Types.FLOAT:
                case Types.REAL:
                case Types.DOUBLE:
                case Types.NUMERIC:
                case Types.DECIMAL:
                    addDoubleValue(sheet, i, rs);
                    break;
                case Types.CHAR:
                case Types.VARCHAR:
                case Types.LONGVARCHAR:
                    addStringValue(sheet, i, rs);
                    break;
                case Types.DATE:
                case Types.TIME:
                case Types.TIMESTAMP:
                    addDateValue(sheet, i, rs);
            }
        } catch (Exception ex) {
            // Want to catch unchecked exceptions.
            LOGGER.error("Error converting cell", ex);
            sheet.addCell("null");
        }
    }

    /**
     * Method to add a double value to a worksheet
     * @param sheet The worksheet
     * @param i The column index
     * @param rs Result set containing the data
     * @throws SQLException
     */
    private static void addDoubleValue(MyWorksheet sheet, int i, ResultSet rs) throws SQLException {
        double doubleValue = rs.getDouble(i + 1);
        if (!rs.wasNull()) {
            sheet.addCell(doubleValue);
        } else {
            sheet.addCell("null");
        }
    }

    public static void BuildSpreadsheetFromQuery(DataSource dataSource, String query, OutputStream out) {
        Connection conn = null;
        Statement stmt = null;
        ResultSet rs = null;
        ResultSetMetaData rsmd;
        MyWorkbook wb = null;
        MyWorksheet sheet = null;
        try {
            conn = dataSource.getConnection();
            stmt = conn.createStatement();
            rs = stmt.executeQuery(query);
            rsmd = rs.getMetaData();
            int numColumns = rsmd.getColumnCount();
            String[] columnNames = new String[numColumns];
            int[] columnTypes = new int[numColumns];
            for (int i = 0; i < numColumns; i++) {
                columnNames[i] = rsmd.getColumnName(i + 1);
                columnTypes[i] = rsmd.getColumnType(i + 1);
            }
            wb = new MyWorkbook(out);
            sheet = wb.getWorksheet();
            sheet.startRow();
            // for each column in the query results, create a spreadsheet column
            for (int i = 0; i < numColumns; i++) {
                sheet.addCell(columnNames[i]);
            }
            sheet.endRow();
            // For each row in the query results, create a spreadsheet row
            while (rs.next()) {
                sheet.startRow();
                for (int i = 0; i < numColumns; i++) {
                    addColumn(columnTypes[i], i, sheet, rs);
                }
                sheet.endRow();
            }
            sheet.close();
            sheet = null;
            wb.close();
            wb = null;
        } catch (SQLException ex) {
            LOGGER.error("Error reading table", ex);
        } catch (Throwable ex) {
            // Catch any unexcpetied condition
            LOGGER.error("Unexpected fatal condition", ex);
        } finally {
            if (sheet != null) {
                sheet.close();
            }
            if (wb != null) {
                wb.close();
            }
            try {
                if (rs != null) {
                    rs.close();
                }
            } catch (SQLException e) {
                /* nothing more can be done */
            }
            try {
                if (stmt != null) {
                    stmt.close();
                }
            } catch (SQLException e) {
                /* nothing more can be done */
            }
            try {
                if (conn != null) {
                    conn.close();
                }
            } catch (SQLException e) {
                /* nothing more can be done */
            }
        }
    }

    /**
     * Method to add a date value to a worksheet
     * @param sheet The worksheet
     * @param i The column index
     * @param rs Result set containing the data
     * @throws SQLException
     */
    private static void addDateValue(MyWorksheet sheet, int i, ResultSet rs) throws SQLException {
        String dateString = rs.getString(i + 1);
        if (!rs.wasNull()) {
            sheet.addDateCell(dateString);
        } else {
            sheet.addCell("null");
        }
    }

    /**
     * Method to add a string value to a worksheet
     * @param sheet The worksheet
     * @param i The column index
     * @param rs Result set containing the data
     * @throws SQLException
     */
    private static void addStringValue(MyWorksheet sheet, int i, ResultSet rs) throws SQLException {
        String stringValue = rs.getString(i + 1);
        if (!rs.wasNull()) {
            sheet.addCell(stringValue);
        } else {
            sheet.addCell("null");
        }
    }
    
}
