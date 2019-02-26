package org.jxls.demo.model;


import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Sample Department bean to demostrate main excel export features
 * @author Leonid Vysochyn
 */
public class Department {
    private String name;
    private Employee chief = new Employee();
    private List<Employee> staff = new ArrayList<Employee>();
    private String link;
    private byte[] image;
    private List<Employee> staff2 = new ArrayList<>();


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

    public static List<Department> generate(int depCount, int employeeCount){
        List<Department> departments = new ArrayList<Department>();
        Random random = new Random(System.currentTimeMillis());
        for(int index = 0; index < depCount; index++){
            Department dep = new Department("Dep " + index);
            dep.setChief( Employee.generateOne("ch" + index));
            dep.setStaff( Employee.generate(1 + random.nextInt(employeeCount)) );
            departments.add( dep );
        }
        return departments;
    }

    public static List<Department> generate(int depCount, int employeeCount, int otherEmployeeCount){
        List<Department> departments = new ArrayList<Department>();
        Random random = new Random(System.currentTimeMillis());
        for(int index = 0; index < depCount; index++){
            Department dep = new Department("Dep " + index);
            dep.setChief( Employee.generateOne("ch" + index));
            dep.setStaff( Employee.generate(1 + random.nextInt(employeeCount)) );
            dep.setStaff2( Employee.generate(1 + random.nextInt(otherEmployeeCount)) );
            departments.add( dep );
        }
        return departments;
    }

    public void addEmployee(Employee employee) {
        staff.add(employee);
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

    public void setStaff(List staff) {
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

    @Override
    public String toString() {
        return "Department{" +
                "name='" + name + '\'' +
                ", chief=" + chief +
                ", staff.size=" + staff.size() +
                '}';
    }
}
