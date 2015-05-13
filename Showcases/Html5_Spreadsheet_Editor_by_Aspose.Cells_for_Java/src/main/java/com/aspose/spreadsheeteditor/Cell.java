package com.aspose.spreadsheeteditor;

import java.io.Serializable;
import java.util.HashSet;
import java.util.StringJoiner;

public class Cell implements Serializable {

    private String name;
    private String value;
    private String formula;
    private String style;
    private HashSet<String> cssClass = new HashSet<>(); //NOSONAR
    private int columnId;
    private int rowId;
    private int colspan = 0;
    private int rowspan = 0;
    private boolean bold;
    private boolean italic;
    private boolean underline;

    public String getName() {
        return this.name;
    }

    public Cell setName(String name) {
        this.name = name;
        return this;
    }

    public Cell setValue(String value) {
        this.value = value;
        return this;
    }

    public Cell setFormula(String formula) {
        this.formula = formula;
        return this;
    }

    public String getValue() {
        return value;
    }

    public String getFormula() {
        return formula;
    }

    public String getStyle() {
        return style;
    }

    public Cell addClass(String c) {
        this.cssClass.add(c);
        return this;
    }

    public Cell removeClass(String c) {
        this.cssClass.remove(c);
        return this;
    }

    public String getCssClass() {
        StringJoiner j = new StringJoiner(" ");
        for (String s : this.cssClass) {
            j.add(s);
        }
        return j.toString();
    }

    public Cell setStyle(String style) {
        this.style = style;
        return this;
    }

    public int getColumnId() {
        return columnId;
    }

    public Cell setColumnId(int columnId) {
        this.columnId = columnId;
        return this;
    }

    public int getRowId() {
        return rowId;
    }

    public Cell setRowId(int rowId) {
        this.rowId = rowId;
        return this;
    }

    public int getColspan() {
        return colspan;
    }

    public Cell setColspan(int colspan) {
        this.colspan = colspan;
        return this;
    }

    public int getRowspan() {
        return rowspan;
    }

    public Cell setRowspan(int rowspan) {
        this.rowspan = rowspan;
        return this;
    }

    public boolean isBold() {
        return bold;
    }

    public Cell setBold(boolean bold) {
        this.bold = bold;
        return this;
    }

    public boolean isItalic() {
        return italic;
    }

    public Cell setItalic(boolean italic) {
        this.italic = italic;
        return this;
    }

    public boolean isUnderline() {
        return underline;
    }

    public Cell setUnderline(boolean underline) {
        this.underline = underline;
        return this;
    }

}
