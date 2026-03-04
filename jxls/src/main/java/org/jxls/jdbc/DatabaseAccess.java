package org.jxls.jdbc;

import static java.sql.Types.CLOB;

import java.io.IOException;
import java.io.Reader;
import java.sql.Connection;
import java.sql.ParameterMetaData;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.JxlsException;

/**
 * Execute SQL queries
 */
public class DatabaseAccess {
    private final Connection conn;
    private final Integer fetchSize;

    public DatabaseAccess(Connection connection) {
        this(connection, null);
    }
    
    public DatabaseAccess(Connection connection, Integer fetchSize) {
        if (connection == null) {
            throw new IllegalArgumentException("connection must not be null");
        }
        conn = connection;
        this.fetchSize = fetchSize;
    }

    public List<Map<String, Object>> query(final String sql, Object... params) {
        if (sql == null) {
            throw new IllegalArgumentException("SQL statement must not be null");
        }
        try (PreparedStatement stmt = conn.prepareStatement(sql.replace(XlsCommentAreaBuilder.LINE_SEPARATOR, System.lineSeparator()))) {
            if (fetchSize != null) {
                stmt.setFetchSize(fetchSize.intValue());
            }
            fillStatement(stmt, params);
            try (ResultSet rs = stmt.executeQuery()) {
                return handle(rs);
            }
        } catch (Exception e) {
            throw new JxlsException("Failed to execute SQL query: " + e.getMessage(), e);
        }
    }

    // The implementation is a slightly modified version of a similar method of AbstractQueryRunner in Apache DBUtils
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
        } catch (Exception e) {
            pmdKnownBroken = true;
        }

        if (stmtCount != paramsCount) {
            throw new JxlsException("SQL statement error. Wrong number of parameters: expected "
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

    protected List<Map<String, Object>> handle(ResultSet rs) throws SQLException {
        List<Map<String, Object>> rows = new ArrayList<>();
        while (rs.next()) {
            rows.add(handleRow(rs));
        }
        return rows;
    }

    private Map<String, Object> handleRow(ResultSet rs) throws SQLException {
        Map<String, Object> result = new CaseInsensitiveHashMap();
        ResultSetMetaData meta = rs.getMetaData();
        int cols = meta.getColumnCount();
        for (int col = 1; col <= cols; col++) {
            String columnName = meta.getColumnLabel(col);
            if (columnName == null || columnName.isEmpty()) {
                columnName = meta.getColumnName(col);
            }
            if (meta.getColumnType(col) == CLOB) {
                try {
                    java.sql.Clob clob = rs.getClob(col);
                    if (clob == null) {
                        result.put(columnName, "");
                        continue;
                    }
                    Reader inStream = clob.getCharacterStream();
                    char[] c = new char[(int) clob.length()];
                    inStream.read(c);
                    String data = new String(c);
                    inStream.close();
                    result.put(columnName, data);
                } catch (IOException e) {
                    throw new JxlsException("Error reading CLOB field " + columnName, e);
                }
            } else {
                result.put(columnName, rs.getObject(col));
            }
        }
        return result;
    }
}
