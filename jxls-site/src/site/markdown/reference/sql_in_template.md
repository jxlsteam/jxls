Using SQL in Template
=====================

Introduction
------------
Jxls allows you to use SQL queries directly in Excel template to produce a collection 
that can be processed by [Each-Command](each_command.html).
The idea is to use an utility class that executes the SQL statement and converts the result set into a list of objects.
Jxls ships with *JdbcHelper* class which can be used for this purpose.  

JdbcHelper
---------------------
To execute SQL queries in the template you can put an instance of *JdbcHelper* class into the context.    
*JdbcHelper* object can be constructed by passing JDBC [Connection](http://docs.oracle.com/javase/7/docs/api/java/sql/Connection.html) instance to its constructor.
 
    Connection conn = ... // get JDBC connection
    JdbcHelper jdbcHelper = new JdbcHelper(conn);
    context.putVar("jdbc", jdbcHelper);
    
Next you can refer to *jdbc* object in your Excel template to execute SQL queries for example
    
    jx:each(items="jdbc.query('select * from employee where payment > ?', 2000)" var="employee" lastCell="C4")
    
In the above command we passed `select * from employee where payment > ?` SQL statement and also specified the substitution parameter (2000) 
to *query* method of the *JdbcHelper* instance.The *query* method has the following signature
      
      public List<Map<String, Object>> query(String sql, Object... params)
      
The method takes a query string and a list of substitution parameters to the query. During execution the method uses the [Connection](http://docs.oracle.com/javase/7/docs/api/java/sql/Connection.html)
to execute the passed SQL using [PreparedStatement](http://docs.oracle.com/javase/7/docs/api/java/sql/PreparedStatement.html).

Next it converts each row of the [ResultSet](http://docs.oracle.com/javase/7/docs/api/java/sql/ResultSet.html) to *Map<String, Object>*
where keys of the map are the case insensitive column names and the values are the corresponding column values.
      
The created *List<Map<String, Object>>* is then processed by [jx:each command](each_command.html) in a regular way.
      
See [SqlDemo Sample](../samples/sql_demo.html) to see a working example of SQL usage in the template.      