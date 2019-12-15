package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.invoke.MethodHandles;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.jdbc.JdbcHelper;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn 17-Oct-2015
 */
public class SqlDemo {
    private static final Logger logger = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass());
    public static final String DRIVER = "org.apache.derby.jdbc.EmbeddedDriver";
    public static final String CONNECTION_URL = "jdbc:derby:memory:testDB;create=true";

    @Test
    public void test() throws ParseException, IOException, ClassNotFoundException, SQLException {
        logger.info("Running SQL demo");
        Class.forName(DRIVER);
        try (Connection conn = DriverManager.getConnection(CONNECTION_URL)) {
            initData(conn);
            JdbcHelper jdbcHelper = new JdbcHelper(conn);
            try (InputStream is = SqlDemo.class.getResourceAsStream("sql_demo_template.xls")) {
                try (OutputStream os = new FileOutputStream("target/sql_demo_output.xls")) {
                    Context context = new Context();
                    context.putVar("conn", conn);
                    context.putVar("jdbc", jdbcHelper);
                    JxlsHelper.getInstance().processTemplate(is, os, context);
                }
            }
        }
    }

    private static void initData(Connection conn) throws SQLException, ParseException {
        String createTableSlq = "CREATE TABLE employee (" +
                "id INT NOT NULL, " +
                "name VARCHAR(20) NOT NULL, " +
                "birthdate DATE, " +
                "payment DECIMAL, " +
                "PRIMARY KEY (id))";
        String insertSql = "INSERT INTO employee VALUES (?,?,?,?)";
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (Statement stmt = conn.createStatement()) {
            stmt.executeUpdate(createTableSlq);
            int k = 1;
            try (PreparedStatement insertStmt = conn.prepareStatement(insertSql)) {
                for (Employee employee : employees) {
                    insertStmt.setInt(1, k++);
                    insertStmt.setString(2, employee.getName());
                    insertStmt.setDate(3, new Date(employee.getBirthDate().getTime()));
                    insertStmt.setBigDecimal(4, employee.getPayment());
                    insertStmt.executeUpdate();
                }
            }
        }
    }
}
