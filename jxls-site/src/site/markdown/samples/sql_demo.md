SQL Demo Example
====================================

Introduction
-------------

This example shows how to use database queries in your template.
For reference documentation about using SQL queries in Excel template see [SQL in Template](../reference/sql_in_template.html)

Excel template
---------------

The [excel template](../xls/sql_demo_template.xls) for this example looks like this

![SQL Demo template](../images/sql_demo_template.png)

Note that we have SQL statement in *items* attribute of *jx:each* command

    jx:each(items="jdbc.query('select * from employee where payment > ?', 2000)" var="employee" lastCell="C4")

Java code
---------

The example uses Apache Derby in memory database and creates a sample EMPLOYEE table in the *initData()* method and then inserts some
Employee data into the table

        private static void initData(Connection conn) throws SQLException, ParseException {
            String createTableSlq = "CREATE TABLE employee (" +
                    "id INT NOT NULL, " +
                    "name VARCHAR(20) NOT NULL, " +
                    "birthdate DATE, " +
                    "payment DECIMAL, " +
                    "PRIMARY KEY (id))";
            String insertSql = "INSERT INTO employee VALUES (?,?,?,?)";
            List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
                try(Statement stmt = conn.createStatement()){
                    stmt.executeUpdate(createTableSlq);
                    int k = 1;
                    try(PreparedStatement insertStmt = conn.prepareStatement(insertSql)){
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

The *main* method is shown below

        public static void main(String[] args) throws ParseException, IOException, ClassNotFoundException, SQLException {
            logger.info("Running SQL demo");
            Class.forName(DRIVER);
            try (Connection conn = DriverManager.getConnection(CONNECTION_URL)){
                initData(conn);
                JdbcHelper jdbcHelper = new JdbcHelper(conn);
                try(InputStream is = SqlDemo.class.getResourceAsStream("sql_demo_template.xls")) {
                    try (OutputStream os = new FileOutputStream("target/sql_demo_output.xls")) {
                        Context context = new Context();
                        context.putVar("conn", conn);
                        context.putVar("jdbc", jdbcHelper);
                        JxlsHelper.getInstance().processTemplate(is, os, context);
                    }
                }
            }
        }

Here we first create *JdbcHelper* instance

    JdbcHelper jdbcHelper = new JdbcHelper(conn);

and then put it into the context and invoke the template processing

    Context context = new Context();
    context.putVar("conn", conn);
    context.putVar("jdbc", jdbcHelper);
    JxlsHelper.getInstance().processTemplate(is, os, context);


Excel output
------------

Final [report](../xls/sql_demo_output.xls) for this example is shown on the following screenshot

![SQL Demo output](../images/sql_demo_output.png)