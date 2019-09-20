package org.jxls.demo.guide;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author Leonid Vysochyn
 */
public class Employee {
    private String name;
    private Date birthDate;
    private BigDecimal payment;
    private BigDecimal bonus;

    private String buGroup;

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

}
