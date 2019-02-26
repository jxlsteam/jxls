package org.jxls.demo.model;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * Sample Employee bean to demonstrate simple export features
 * @author Leonid Vysochyn
 */
public class Employee {
    private String name;
    private int age;
    private Double payment;
    private Double bonus;
    private Date birthDate;
    private Employee superior;

    static Random random = new Random(System.currentTimeMillis());
    static long current = System.currentTimeMillis();

    public Employee() {
    }

    public Employee(String name, int age, Double payment, Double bonus) {
        this.name = name;
        this.age = age;
        this.payment = payment;
        this.bonus = bonus;
    }

    public Employee(String name, int age, double payment, double bonus, Date birthDate) {
        this.name = name;
        this.age = age;
        this.payment = payment;
        this.bonus = bonus;
        this.birthDate = birthDate;
    }

    public Employee(String name, int age, double payment, double bonus) {
        this.name = name;
        this.age = age;
        this.payment = payment;
        this.bonus = bonus;
    }

    public Employee(String name, Date birthDate, Double payment, Double bonus) {
        this.name = name;
        this.birthDate = birthDate;
        this.payment = payment;
        this.bonus = bonus;
    }


    public static List<Employee> generate(int num){
        List<Employee> result = new ArrayList<Employee>();
        for(int index = 0; index < num; index++){
            result.add( generateOne("" + index) );
        }
        return result;
    }

    public static Employee generateOne(String nameSuffix){
        return new Employee("Employee " + nameSuffix, random.nextInt(100), 1000 + random.nextDouble()*5000, random.nextInt(100)/100.0d, new Date(current - (1000000 + random.nextInt(1000000))));
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Double getPayment() {
        return payment;
    }

    public void setPayment(Double payment) {
        this.payment = payment;
    }

    public Double getBonus() {
        return bonus;
    }

    public void setBonus(Double bonus) {
        this.bonus = bonus;
    }

    public Date getBirthDate() {
        return birthDate;
    }

    public void setBirthDate(Date birthDate) {
        this.birthDate = birthDate;
    }

    public Employee getSuperior() {
        return superior;
    }

    public void setSuperior(Employee superior) {
        this.superior = superior;
    }

    @Override
    public String toString() {
        return "Employee{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", payment=" + payment +
                '}';
    }
}
