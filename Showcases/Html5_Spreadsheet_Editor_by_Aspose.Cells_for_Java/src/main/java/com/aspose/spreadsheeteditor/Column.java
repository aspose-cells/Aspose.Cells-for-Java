package com.aspose.spreadsheeteditor;

import java.io.Serializable;

/**
 *
 * @author saqib
 */
public class Column implements Serializable {

    private int id;
    private String name;
    private String header;
    private String property;
    private int width;
    
    public Column(int id, String name) {
        this.id = id;
        this.name = this.header = this.property = name;
    }

    public String getHeader() {
        return header;
    }

    public String getProperty() {
        return property;
    }

    public int getId() {
        return id;
    }

    public String getName() {
        return name;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }
}

