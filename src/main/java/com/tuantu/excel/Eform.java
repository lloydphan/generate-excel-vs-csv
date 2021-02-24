package com.tuantu.excel;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class Eform {

    private BigInteger id;

    private String name;

    private String contact;

    private String email;

    private List<String> listCols;

    public Eform(BigInteger id, String name, String contact, String email, String colVal, long limit) {
        this.id = id;
        this.name = name;
        this.contact = contact;
        this.email = email;
        listCols = new ArrayList<>();
        for (int i = 1; i <= limit; i++) {
            listCols.add(colVal + i);
        }
    }

    public BigInteger getId() {
        return id;
    }

    public void setId(BigInteger id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getContact() {
        return contact;
    }

    public void setContact(String contact) {
        this.contact = contact;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public List<String> getListCols() {
        return listCols;
    }

    public void setListCols(List<String> listCols) {
        this.listCols = listCols;
    }
}
