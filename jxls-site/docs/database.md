# Database access

Jxls allows you to use SQL queries directly in the Excel template to produce a collection 
that can be processed by the [jx:each command](each.html).
The idea is to use an utility class that executes the SQL statement and converts the result set into a list of objects.
Jxls ships with `DatabaseAccess` class which can be used for this purpose.  

## DatabaseAccess

To execute SQL queries in the template you can put an instance of *DatabaseAccess* class into the data map.    
*DatabaseAccess* object can be constructed by passing JDBC Connection instance to its constructor.
You can also specify the fetch size as a 2nd argument.

```
Connection conn = ... // get JDBC connection
DatabaseAccess db = new DatabaseAccess(conn);
data.put("jdbc", db);
```
    
Use jdbc like this in your jx:each command:

```
jx:each(items="jdbc.query('select * from employee where payment > ?', 2000)" var="employee" lastCell="C4")
```
    
In the above command we passed `select * from employee where payment > ?` SQL statement and also specified the substitution parameter (2000) 
to the query() method of the DatabaseAccess object. The query() method has this signature:

```
public List<Map<String, Object>> query(String sql, Object... params)
```
      
The method takes a query string and a list of substitution parameters to the query. During execution the method uses the Connection
to execute the passed SQL using PreparedStatement.
Next it converts each row of the ResultSet to `Map<String, Object>`
where keys of the map are the case insensitive column names and the values are the corresponding column values.
The created `List<Map<String, Object>>` is then processed by the jx:each command in a regular way.

Single quotation marks in SQL are to be escaped with backslash:

```
jx:each(items="jdbc.query('select * from employee where name=\'Elsa\' ')" var="employee" lastCell="C4")
```

Multi-line SQL is supported. This feature can be turned off using `XlsCommentAreaBuilder.MULTI_LINE_SQL_FEATURE=false`.
