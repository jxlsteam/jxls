package org.jxls.entity;

import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Random;

public class Employee {
    private static final Random random = new Random(System.currentTimeMillis());
    private static long current = System.currentTimeMillis();
    private String name;
    private Date birthDate;
    private BigDecimal payment;
    private BigDecimal bonus;
    private String buGroup;
    private int age;
    private Department departmentObject;

    public Employee(String name, Date birthDate, BigDecimal payment, BigDecimal bonus) {
        this.name = name;
        this.birthDate = birthDate;
        this.payment = payment;
        this.bonus = bonus;
    }

    public Employee(String name, Date birthDate, double payment, double bonus) {
        this(name, birthDate, new BigDecimal(payment), new BigDecimal(bonus));
    }

    public Employee(String name, Date birthDate, double payment, double bonus, String buGroup) {
        this(name, birthDate, new BigDecimal(payment), new BigDecimal(bonus), buGroup);
    }

    public Employee(String name, Date birthDate, BigDecimal payment, BigDecimal bonus, String buGroup) {
        this.name = name;
        this.birthDate = birthDate;
        this.payment = payment;
        this.bonus = bonus;
        this.buGroup = buGroup;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getBirthDate() {
        return birthDate;
    }

    public void setBirthDate(Date birthDate) {
        this.birthDate = birthDate;
    }

    public BigDecimal getPayment() {
        return payment;
    }

    public void setPayment(BigDecimal payment) {
        this.payment = payment;
    }

    public BigDecimal getBonus() {
        return bonus;
    }

    public void setBonus(BigDecimal bonus) {
        this.bonus = bonus;
    }

    public String getBuGroup() {
        return buGroup;
    }

    public void setBuGroup(String buGroup) {
        this.buGroup = buGroup;
    }

    public static List<Employee> generateSampleEmployeeData() {
        try {
            List<Employee> employees = new ArrayList<>();
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
            employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
            employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25) );
            employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00) );
            employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15) );
            employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20) );
            return employees;
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }

    public static List<Employee> generate(int num) {
        List<Employee> result = new ArrayList<>();
        for (int index = 0; index < num; index++) {
            result.add(generateOne("" + index));
        }
        return result;
    }

    public static Employee generateOne(String nameSuffix) {
        Employee ret = new Employee("Employee " + nameSuffix,
                new Date(current - (1000000 + random.nextInt(1000000))),
                1000 + random.nextDouble() * 5000,
                random.nextInt(100) / 100.0d);
        ret.setAge(random.nextInt(100));
        return ret;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Department getDepartmentObject() {
        return departmentObject;
    }
    
    public Employee withDepartmentKey(String key) {
        departmentObject = new Department();
        departmentObject.setKey(key);
        return this;
    }
}
