package com.tkachenko.buyerhelper.entity;

import javax.persistence.*;

@Entity
@Table(name = "files_table")
public class FileEntity {
    @Id
    @Column(name = "id")
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;

    @Column(name = "original_file_name")
    private String originalFileName;

    @Column(name = "storage_file_name")
    private String storageFileName;

    @Column(name = "year")
    private String year;

    @Column(name = "month")
    private String month;

    @Column(name = "day")
    private String day;

    @Column(name = "time")
    private String time;

    @Column(name = "is_actual")
    private boolean isActual;

    public FileEntity(String originalFileName, String storageFileName, String year, String month, String day, String time, boolean isActual) {
        this.originalFileName = originalFileName;
        this.storageFileName = storageFileName;
        this.year = year;
        this.month = month;
        this.day = day;
        this.time = time;
        this.isActual = isActual;
    }

    public FileEntity() {
    }

    public Long getId() {
        return id;
    }

    public String getOriginalFileName() {
        return originalFileName;
    }

    public void setOriginalFileName(String originalFileName) {
        this.originalFileName = originalFileName;
    }

    public String getStorageFileName() {
        return storageFileName;
    }

    public void setStorageFileName(String storageFileName) {
        this.storageFileName = storageFileName;
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public String getDay() {
        return day;
    }

    public void setDay(String day) {
        this.day = day;
    }

    public String getTime() {
        return time;
    }

    public void setTime(String time) {
        this.time = time;
    }

    public boolean isActual() {
        return isActual;
    }

    public void setActual(boolean actual) {
        isActual = actual;
    }

    @Override
    public String toString() {
        return "FileEntity{" +
                "id=" + id +
                ", originalFileName='" + originalFileName + '\'' +
                ", storageFileName='" + storageFileName + '\'' +
                ", year='" + year + '\'' +
                ", month='" + month + '\'' +
                ", day='" + day + '\'' +
                ", time='" + time + '\'' +
                ", isActual=" + isActual +
                '}';
    }
}
