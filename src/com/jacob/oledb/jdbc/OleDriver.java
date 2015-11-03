package com.jacob.oledb.jdbc;

import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverPropertyInfo;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.SQLFeatureNotSupportedException;
import java.util.Properties;
import java.util.logging.Logger;

public class OleDriver implements Driver {

    public static void main(String[] args) {
        try {
            Driver driver = (Driver) Class.forName("com.jacob.oledb.jdbc.OleDriver").newInstance();
            Connection c = driver.connect("jdbc:oledb:Provider=vfpoledb;Data Source=C:\\Siaf_Vfp\\Data;Collating Sequence=general", null);
            ResultSet rs = c.createStatement().executeQuery("SELECT * FROM meta");
            ResultSetMetaData m = rs.getMetaData();
            int nc = m.getColumnCount();
            for (int i = 1; i <= nc; i++) {
                System.out.print(m.getColumnName(i) + "\t");
            }
            System.out.println();
            while (rs.next()) {
                for (int i = 1; i <= nc; i++) {
                    System.out.print(rs.getString(i) + "\t");
                }
                System.out.println();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public Connection connect(String url, Properties info) throws SQLException {
        OleConnection c = new OleConnection();
        url = url.substring("jdbc:oledb:".length());
        c.setConnectionString(url);
        c.Open();
        return c;
    }

    @Override
    public boolean acceptsURL(String url) throws SQLException {
        return url.startsWith("jdbc:oledb:");
    }

    @Override
    public DriverPropertyInfo[] getPropertyInfo(String url, Properties info) throws SQLException {
        return null;
    }

    @Override
    public int getMajorVersion() {
        return 0;
    }

    @Override
    public int getMinorVersion() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean jdbcCompliant() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public Logger getParentLogger() throws SQLFeatureNotSupportedException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

}
