package org.jxls.examples;

import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.entity.Employee;
import org.jxls.jdbc.DatabaseAccess;

public class SqlDemo {
    public static final String DRIVER = "org.h2.Driver";
    public static final String CONNECTION_URL = "jdbc:h2:mem:testDB;DB_CLOSE_DELAY=-1";

    @Test
    public void test() throws ClassNotFoundException, SQLException {
        try (Connection conn = openConnection()) {
            Context context = new ContextImpl();
            context.putVar("jdbc", new DatabaseAccess(conn));
            
            JxlsTester tester = JxlsTester.xlsx(getClass());
            tester.processTemplate(context);
        }
    }
    
    private Connection openConnection() throws ClassNotFoundException, SQLException {
        Connection conn = DriverManager.getConnection(CONNECTION_URL);
        try (Statement stmt = conn.createStatement()) {
            String createTableSlq = "CREATE TABLE employee (" +
                    "id INT NOT NULL, " +
                    "name VARCHAR(20) NOT NULL, " +
                    "birthdate DATE, " +
                    "payment DECIMAL, " +
                    "PRIMARY KEY (id))";
            stmt.executeUpdate(createTableSlq);
            int k = 1;
            String insertSql = "INSERT INTO employee VALUES (?,?,?,?)";
            try (PreparedStatement insertStmt = conn.prepareStatement(insertSql)) {
                List<Employee> employees = Employee.generateSampleEmployeeData();
                for (Employee employee : employees) {
                    insertStmt.setInt(1, k++);
                    insertStmt.setString(2, employee.getName());
                    insertStmt.setDate(3, new Date(employee.getBirthDate().getTime()));
                    insertStmt.setBigDecimal(4, employee.getPayment());
                    insertStmt.executeUpdate();
                }
            }
        }
        return conn;
    }
}
