package org.jxls.jdbc;

import org.jxls.common.JxlsException;

import java.sql.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Leonid Vysochyn on 17-Oct-15.
 */
public class JdbcHelper {
    Connection conn;

    public JdbcHelper(Connection conn) {
        this.conn = conn;
    }

    public List<Map<String, Object>> query(String sql, Object... params)  {
        List<Map<String, Object>> result;
        if (conn == null) {
            throw new JxlsException("Null jdbc connection");
        }

        if (sql == null) {
            throw new JxlsException("Null SQL statement");
        }

        try(PreparedStatement stmt = conn.prepareStatement(sql)){
            fillStatement(stmt, params);
            try(ResultSet rs = stmt.executeQuery()){
                result = handle(rs);
            }
        } catch (Exception e) {
            throw new JxlsException("Failed to execute sql", e);
        }

        return result;
    }

    /*
     * The implementation is a slightly modified version of a similar method of AbstractQueryRunner in Apache DBUtils
     */
    private void fillStatement(PreparedStatement stmt, Object[] params) throws SQLException {
        // nothing to do here
        if (params == null) {
            return;
        }

        // check the parameter count, if we can
        ParameterMetaData pmd = null;
        boolean pmdKnownBroken = false;

        int stmtCount = 0;
        int paramsCount = 0;
        try {
            pmd = stmt.getParameterMetaData();
            stmtCount = pmd.getParameterCount();
            paramsCount = params.length;
        }catch (Exception e){
            pmdKnownBroken = true;
        }

        if (stmtCount != paramsCount) {
            throw new SQLException("Wrong number of parameters: expected "
                    + stmtCount + ", was given " + paramsCount);
        }

        for (int i = 0; i < params.length; i++) {
            if (params[i] != null) {
                stmt.setObject(i + 1, params[i]);
            } else {
                // VARCHAR works with many drivers regardless
                // of the actual column type. Oddly, NULL and
                // OTHER don't work with Oracle's drivers.
                int sqlType = Types.VARCHAR;
                if (!pmdKnownBroken) {
                    try {
                        /*
                         * It's not possible for pmdKnownBroken to change from
                         * true to false, (once true, always true) so pmd cannot
                         * be null here.
                         */
                        sqlType = pmd.getParameterType(i + 1);
                    } catch (SQLException e) {
                        pmdKnownBroken = true;
                    }
                }
                stmt.setNull(i + 1, sqlType);
            }
        }
    }

    private List<Map<String, Object>> handle(ResultSet rs) throws SQLException {
        List<Map<String,Object>> rows = new ArrayList<>();
        while (rs.next()) {
            rows.add(handleRow(rs));
        }
        return rows;
    }

    private Map<String, Object> handleRow(ResultSet rs) throws SQLException {
        Map<String, Object> result = new LinkedHashMap<>();
        ResultSetMetaData rsmd = rs.getMetaData();
        int cols = rsmd.getColumnCount();
        for (int i = 1; i <= cols; i++) {
            String columnName = rsmd.getColumnLabel(i);
            if (null == columnName || 0 == columnName.length()) {
                columnName = rsmd.getColumnName(i);
            }
            result.put(columnName, rs.getObject(i));
        }
        return result;
    }


}
