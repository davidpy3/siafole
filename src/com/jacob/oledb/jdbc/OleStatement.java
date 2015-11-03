package com.jacob.oledb.jdbc;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.SQLWarning;
import java.sql.Statement;

/**
 * Custom dispatch object to make it easy for us to provide application specific
 * API.
 *
 */
public class OleStatement extends Dispatch implements Statement {

    private OleConnection oleConnection;

    OleStatement(OleConnection c) {
        oleConnection=c;
    }

    @Override
    public ResultSet executeQuery(String sql) throws SQLException {
        OleStatement o=new OleStatement();
        o.setActiveConnection(oleConnection);
        o.setCommandType(CommandTypeEnum.adCmdText);
        o.setCommandText(sql);
        return o.Execute();
    }

    /**
     * standard constructor
     */
    public OleStatement() {
        super("ADODB.Command");
    }

    /**
     * This constructor is used instead of a case operation to turn a Dispatch
     * object into a wider object - it must exist in every wrapper class whose
     * instances may be returned from method calls wrapped in VT_DISPATCH
     * Variants.
     *
     * @param dispatchTarget
     */
    public OleStatement(Dispatch dispatchTarget) {
        super(dispatchTarget);
    }

    /**
     * runs the "Properties" command
     *
     * @return the properties
     */
    public Variant getProperties() {
        return Dispatch.get(this, "Properties");
    }

    /**
     * runs the "ActiveConnection" command
     *
     * @return a Connection object
     */
    public OleConnection getActiveConnection() {
        return new OleConnection(Dispatch.get(this, "ActiveConnection")
                .toDispatch());
    }

    /**
     * Sets the "ActiveConnection" object
     *
     * @param ppvObject the new connection
     */
    public void setActiveConnection(OleConnection ppvObject) {
        Dispatch.put(this, "ActiveConnection", ppvObject);
    }

    /**
     *
     * @return the results from "CommandText"
     */
    public String getCommandText() {
        return Dispatch.get(this, "CommandText").toString();
    }

    /**
     *
     * @param pbstr the new "CommandText"
     */
    public void setCommandText(String pbstr) {
        Dispatch.put(this, "CommandText", pbstr);
    }

    /**
     *
     * @return the results of "CommandTimeout"
     */
    public int getCommandTimeout() {
        return Dispatch.get(this, "CommandTimeout").getInt();
    }

    /**
     *
     * @param plTimeout the new "CommandTimeout"
     */
    public void setCommandTimeout(int plTimeout) {
        Dispatch.put(this, "CommandTimeout", new Variant(plTimeout));
    }

    /**
     *
     * @return results from "Prepared"
     */
    public boolean getPrepared() {
        return Dispatch.get(this, "Prepared").getBoolean();
    }

    /**
     *
     * @param pfPrepared the new value for "Prepared"
     */
    public void setPrepared(boolean pfPrepared) {
        Dispatch.put(this, "Prepared", new Variant(pfPrepared));
    }

    /**
     * "Execute"s a command
     *
     * @param RecordsAffected
     * @param Parameters
     * @param Options
     * @return
     */
    public OleResultSet Execute(Variant RecordsAffected, Variant Parameters,
            int Options) {
        return (OleResultSet) Dispatch.call(this, "Execute", RecordsAffected,
                Parameters, new Variant(Options)).toDispatch();
    }

    /**
     * "Execute"s a command
     *
     * @return
     */
    public OleResultSet Execute() {
        Variant dummy = new Variant();
        return new OleResultSet(Dispatch.call(this, "Execute", dummy).toDispatch());
    }

    /**
     * creates a parameter
     *
     * @param Name
     * @param Type
     * @param Direction
     * @param Size
     * @param Value
     * @return
     */
    public Variant CreateParameter(String Name, int Type, int Direction,
            int Size, Variant Value) {
        return Dispatch.call(this, "CreateParameter", Name, new Variant(Type),
                new Variant(Direction), new Variant(Size), Value);
    }

    // need to wrap Parameters
    /**
     * @return "Parameters"
     */
    public Variant getParameters() {
        return Dispatch.get(this, "Parameters");
    }

    /**
     *
     * @param plCmdType new "CommandType"
     */
    public void setCommandType(int plCmdType) {
        Dispatch.put(this, "CommandType", new Variant(plCmdType));
    }

    /**
     *
     * @return current "CommandType"
     */
    public int getCommandType() {
        return Dispatch.get(this, "CommandType").getInt();
    }

    /**
     *
     * @return "Name"
     */
    public String getName() {
        return Dispatch.get(this, "Name").toString();
    }

    /**
     *
     * @param pbstrName new "Name"
     */
    public void setName(String pbstrName) {
        Dispatch.put(this, "Name", pbstrName);
    }

    /**
     *
     * @return curent "State"
     */
    public int getState() {
        return Dispatch.get(this, "State").getInt();
    }

    /**
     * cancel whatever it is we're doing
     */
    public void Cancel() {
        Dispatch.call(this, "Cancel");
    }

    @Override
    public int executeUpdate(String sql) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void close() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getMaxFieldSize() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setMaxFieldSize(int max) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getMaxRows() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setMaxRows(int max) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setEscapeProcessing(boolean enable) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getQueryTimeout() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setQueryTimeout(int seconds) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void cancel() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public SQLWarning getWarnings() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void clearWarnings() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setCursorName(String name) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean execute(String sql) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public ResultSet getResultSet() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getUpdateCount() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean getMoreResults() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setFetchDirection(int direction) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getFetchDirection() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setFetchSize(int rows) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getFetchSize() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getResultSetConcurrency() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getResultSetType() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void addBatch(String sql) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void clearBatch() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int[] executeBatch() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public Connection getConnection() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean getMoreResults(int current) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public ResultSet getGeneratedKeys() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int executeUpdate(String sql, int autoGeneratedKeys) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int executeUpdate(String sql, int[] columnIndexes) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int executeUpdate(String sql, String[] columnNames) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean execute(String sql, int autoGeneratedKeys) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean execute(String sql, int[] columnIndexes) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean execute(String sql, String[] columnNames) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public int getResultSetHoldability() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean isClosed() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setPoolable(boolean poolable) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean isPoolable() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void closeOnCompletion() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean isCloseOnCompletion() throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public <T> T unwrap(Class<T> iface) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean isWrapperFor(Class<?> iface) throws SQLException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
}
