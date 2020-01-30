package org.jxls.entity;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Sample Department bean to demonstrate main excel export features
 * 
 * @author Leonid Vysochyn
 */
public class Department {
    private String name;
    private Employee chief = new Employee(null, null, null, null);
    private List<Employee> staff = new ArrayList<>();
    private String link;
    private byte[] image;
    private List<Employee> staff2 = new ArrayList<>();
    private String key;
    
    public Department() {
    }

    public Department(String name) {
        this.name = name;
    }

    public Department(String name, Employee chief, List<Employee> staff) {
        this.name = name;
        this.chief = chief;
        this.staff = staff;
    }

    public int getHeadcount() {
        return staff.size();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getLink() {
        return link;
    }

    public void setLink(String link) {
        this.link = link;
    }

    public Employee getChief() {
        return chief;
    }

    public void setChief(Employee chief) {
        this.chief = chief;
    }

    public List<Employee> getStaff() {
        return staff;
    }

    public void setStaff(List<Employee> staff) {
        this.staff = staff;
    }

    public byte[] getImage() {
        return image;
    }

    public void setImage(byte[] image) {
        this.image = image;
    }

    public List<Employee> getStaff2() {
        return staff2;
    }

    public void setStaff2(List<Employee> staff2) {
        this.staff2 = staff2;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    @Override
    public String toString() {
        return "Department{" + "name='" + name + '\'' + ", chief=" + chief + ", staff.size=" + staff.size() + '}';
    }

    public void addEmployee(Employee employee) {
        staff.add(employee);
    }

    public static List<Department> generate(int depCount, int employeeCount) {
        List<Department> departments = new ArrayList<>();
        Random random = new Random(System.currentTimeMillis());
        for (int index = 0; index < depCount; index++) {
            Department dep = new Department("Dep " + index);
            dep.setChief(Employee.generateOne("ch" + index));
            dep.setStaff(Employee.generate(1 + random.nextInt(employeeCount)));
            departments.add(dep);
        }
        return departments;
    }

    public static List<Department> generate(int depCount, int employeeCount, int otherEmployeeCount) {
        List<Department> departments = new ArrayList<>();
        Random random = new Random(System.currentTimeMillis());
        for (int index = 0; index < depCount; index++) {
            Department dep = new Department("Dep " + index);
            dep.setChief(Employee.generateOne("ch" + index));
            dep.setStaff(Employee.generate(1 + random.nextInt(employeeCount)));
            dep.setStaff2(Employee.generate(1 + random.nextInt(otherEmployeeCount)));
            departments.add(dep);
        }
        return departments;
    }

    public static List<Department> createDepartments() {
        List<Department> departments = new ArrayList<>();
        Department department = new Department("IT");
        Employee chief = createEmployee("Derek", 35, 3000, 0.30);
        department.setChief(chief);
        department.setLink("http://jxls.sf.net");
        department.addEmployee(createEmployee("Elsa", 28, 1500, 0.15));
        department.addEmployee(createEmployee("Oleg", 32, 2300, 0.25));
        department.addEmployee(createEmployee("Neil", 34, 2500, 0.00));
        department.addEmployee(createEmployee("Maria", 34, 1700, 0.15));
        department.addEmployee(createEmployee("John", 35, 2800, 0.20));
        departments.add(department);
        department = new Department("HR");
        chief = createEmployee("Betsy", 37, 2200, 0.30);
        department.setChief(chief);
        department.setLink("http://jxls.sf.net");
        department.addEmployee(createEmployee("Olga", 26, 1400, 0.20));
        department.addEmployee(createEmployee("Helen", 30, 2100, 0.10));
        department.addEmployee(createEmployee("Keith", 24, 1800, 0.15));
        department.addEmployee(createEmployee("Cat", 34, 1900, 0.15));
        departments.add(department);
        department = new Department("BA");
        chief = createEmployee("Wendy", 35, 2900, 0.35);
        department.setChief(chief);
        department.setLink("http://jxls.sf.net");
        department.addEmployee(createEmployee("Denise", 30, 2400, 0.20));
        department.addEmployee(createEmployee("LeAnn", 32, 2200, 0.15));
        department.addEmployee(createEmployee("Natali", 28, 2600, 0.10));
        department.addEmployee(createEmployee("Martha", 33, 2150, 0.25));
        departments.add(department);
        return departments;
    }

    private static Employee createEmployee(String name, int age, double payment, double bonus) {
        Employee ret = new Employee(name, null, payment, bonus);
        ret.setAge(age);
        return ret;
    }
}
