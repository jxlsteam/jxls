package org.jxls.command;

import java.math.BigDecimal;

public class TestEmployee {
    private String department;
    private String name;
    private String job;
    private String city;
    private BigDecimal salary;
    private TestDepartment departmentObject;

    public TestEmployee(String department, String name, String job, String city, double salary) {
        this.department = department;
        this.name = name;
        this.job = job;
        this.city = city;
        this.salary = new BigDecimal(salary);
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getJob() {
        return job;
    }

    public void setJob(String job) {
        this.job = job;
    }

    public String getCity() {
        return city;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public BigDecimal getSalary() {
        return salary;
    }

    public void setSalary(BigDecimal salary) {
        this.salary = salary;
    }
    
    public TestDepartment getDepartmentObject() {
        return departmentObject;
    }
    
    public TestEmployee withDepartmentKey(String key) {
        departmentObject = new TestDepartment(department, key);
        return this;
    }
}
