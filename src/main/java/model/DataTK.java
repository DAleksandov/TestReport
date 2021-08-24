package model;

import java.util.Date;
import java.util.Objects;

public class DataTK {
    private String execution;
    private String  plannedStartDate;
    private String  plannedEndDate;
    private String  actualStartDate;
    private String  actualEndDate;
    private String environment;
    private String testCycle;
    private String keyTK;
    private String nameTK;
    private String status;
    private String owner;
    public DataTK() {

    }

    public void setPlannedStartDate(String plannedStartDate) {
        this.plannedStartDate = plannedStartDate;
    }

    public void setPlannedEndDate(String plannedEndDate) {
        this.plannedEndDate = plannedEndDate;
    }

    public void setActualStartDate(String actualStartDate) {
        this.actualStartDate = actualStartDate;
    }

    public void setActualEndDate(String actualEndDate) {
        this.actualEndDate = actualEndDate;
    }

    public DataTK(String execution, String plannedStartDate, String plannedEndDate, String actualStartDate, String actualEndDate, String environment, String testCycle, String keyTK, String nameTK, String status, String owner) {
        this.execution = execution;
        this.plannedStartDate = plannedStartDate;
        this.plannedEndDate = plannedEndDate;
        this.actualStartDate = actualStartDate;
        this.actualEndDate = actualEndDate;
        this.environment = environment;
        this.testCycle = testCycle;
        this.keyTK = keyTK;
        this.nameTK = nameTK;
        this.status = status;
        this.owner = owner;
    }

    public String getExecution() {
        return execution;
    }

    public void setExecution(String execution) {
        this.execution = execution;
    }

    public String getEnvironment() {
        return environment;
    }

    public void setEnvironment(String environment) {
        this.environment = environment;
    }

    public String getTestCycle() {
        return testCycle;
    }

    public void setTestCycle(String testCycle) {
        this.testCycle = testCycle;
    }

    public String getKeyTK() {
        return keyTK;
    }

    public void setKeyTK(String keyTK) {
        this.keyTK = keyTK;
    }

    public String getNameTK() {
        return nameTK;
    }

    public void setNameTK(String nameTK) {
        this.nameTK = nameTK;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public String getPlannedStartDate() {
        return plannedStartDate;
    }

    public String getPlannedEndDate() {
        return plannedEndDate;
    }

    public String getActualStartDate() {
        return actualStartDate;
    }

    public String getActualEndDate() {
        return actualEndDate;
    }
    
    public String getOwner() {
        return owner;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof DataTK)) return false;
        DataTK dataTK = (DataTK) o;
        return Objects.equals(execution, dataTK.execution) &&
                Objects.equals(plannedStartDate, dataTK.plannedStartDate) &&
                Objects.equals(plannedEndDate, dataTK.plannedEndDate) &&
                Objects.equals(actualStartDate, dataTK.actualStartDate) &&
                Objects.equals(actualEndDate, dataTK.actualEndDate) &&
                Objects.equals(environment, dataTK.environment) &&
                Objects.equals(testCycle, dataTK.testCycle) &&
                Objects.equals(keyTK, dataTK.keyTK) &&
                Objects.equals(nameTK, dataTK.nameTK) &&
                Objects.equals(status, dataTK.status);
    }

    @Override
    public int hashCode() {
        return Objects.hash(execution, plannedStartDate, plannedEndDate, actualStartDate, actualEndDate, environment, testCycle, keyTK, nameTK, status);
    }

    @Override
    public String toString() {
        return "DataTK{" +
                "execution='" + execution + '\'' +
                ", plannedStartDate=" + plannedStartDate +
                ", plannedEndDate=" + plannedEndDate +
                ", actualStartDate=" + actualStartDate +
                ", actualEndDate=" + actualEndDate +
                ", environment='" + environment + '\'' +
                ", testCycle='" + testCycle + '\'' +
                ", keyTK='" + keyTK + '\'' +
                ", nameTK='" + nameTK + '\'' +
                ", status='" + status + '\'' +
                '}';
    }
}
