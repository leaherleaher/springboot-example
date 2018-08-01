package com.springboot.example.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.Date;


@Data
@AllArgsConstructor
public class User {
    private int id;
    private String name;
    private int age;
    private Date date;
}
