package com.aspose.spreadsheeteditor;

import java.awt.Color;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Logger;
import javax.enterprise.context.ApplicationScoped;
import javax.inject.Inject;
import javax.inject.Named;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "cells")
@ApplicationScoped
public class CellsService {

    private static final Logger LOGGER = Logger.getLogger(CellsService.class.getName());

    private HashMap<String, List<Cell>> cells = new HashMap<>(); // NOSONAR
    private HashMap<String, List<Column>> columns = new HashMap<>(); // NOSONAR
    private HashMap<String, List<Row>> rows = new HashMap<>(); // NOSONAR
    private HashMap<String, List<Integer>> columnWidth = new HashMap<>(); // NOSONAR
    private HashMap<String, List<Integer>> rowHeight = new HashMap<>(); // NOSONAR

    @Inject
    private MessageService msg;

    public void put(String key, List<Cell> cells) {
        this.cells.put(key, cells);
    }

    public List<Cell> get(String key) {
        return this.cells.get(key);
    }

    public List<Column> getColumns(String key) {
        return columns.get(key);
    }

    public void putColumns(String key, List<Column> c) {
        columns.put(key, c);
    }

    void removeColumns(String key) {
        columns.remove(key);
    }

    public List<Row> getRows(String key) {
        return rows.get(key);
    }

    public void putRows(String key, List<Row> r) {
        rows.put(key, r);
    }

    void removeRows(String key) {
        rows.remove(key);
    }

    public Cell getCell(String key, int column, int row) {
        return rows.get(key).get(row).getCellsMap().get(com.aspose.cells.CellsHelper.columnIndexToName(column));
    }

    public void putCell(String key, int column, int row, Cell c) {
        rows.get(key).get(row).getCellsMap().put(com.aspose.cells.CellsHelper.columnIndexToName(column), c);
    }

    public List<Integer> getColumnWidth(String key) {
        return columnWidth.get(key);
    }

    public void putColumnWidth(String key, List<Integer> c) {
        columnWidth.put(key, c);
    }

    void removeColumnWidth(String key) {
        columnWidth.remove(key);
    }

    public List<Integer> getRowHeight(String key) {
        return rowHeight.get(key);
    }

    public void putRowHeight(String key, List<Integer> r) {
        rowHeight.put(key, r);
    }

    void removeRowHeight(String key) {
        rowHeight.remove(key);
    }

    public Cell fromBlank(int columnId, int rowId) {
        return new Cell()
                .setColumnId(columnId)
                .setRowId(rowId)
                .setName(com.aspose.cells.CellsHelper.cellIndexToName(rowId, columnId));
    }

    public Cell fromAsposeCell(com.aspose.cells.Cell a) {
        // a = Aspose.Cells' definition of a cell
        // c = Spreassheet's definition of a cell

        Cell c = fromBlank(a.getColumn(), a.getRow());

        try {
            a.calculate(true, null);
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Cell recalculation failure", cx.getMessage());
        }

        try {
            c.setFormula(a.getFormula())
                    .setValue(a.getStringValueWithoutFormat());
        } catch (Exception x) {
            LOGGER.throwing(null, null, x);
            msg.sendMessage("Cell value error", x.getMessage());
        }

        StringBuilder style = new StringBuilder();

        try {
            com.aspose.cells.Color cellFgColor = a.getStyle().getForegroundColor();
            style.append("background-color:")
                    .append(asposeColorToCssColor(cellFgColor, false))
                    .append(";");

            com.aspose.cells.Font font = a.getStyle().getFont();
            style.append("font-family:'").append(font.getName()).append("';");

            if (a.getStyle().getFont().isItalic()) {
                c.addClass("i").setItalic(true);
            }

            if (a.getStyle().getFont().isBold()) {
                c.addClass("b").setBold(true);
            }

            switch (a.getStyle().getFont().getUnderline()) {
                case com.aspose.cells.FontUnderlineType.ACCOUNTING:
                case com.aspose.cells.FontUnderlineType.DASH:
                case com.aspose.cells.FontUnderlineType.DASHED_HEAVY:
                case com.aspose.cells.FontUnderlineType.DASH_DOT_DOT_HEAVY:
                case com.aspose.cells.FontUnderlineType.DASH_DOT_HEAVY:
                case com.aspose.cells.FontUnderlineType.DOTTED:
                case com.aspose.cells.FontUnderlineType.DOTTED_HEAVY:
                case com.aspose.cells.FontUnderlineType.DOT_DASH:
                case com.aspose.cells.FontUnderlineType.DOT_DOT_DASH:
                case com.aspose.cells.FontUnderlineType.DOUBLE:
                case com.aspose.cells.FontUnderlineType.DOUBLE_ACCOUNTING:
                case com.aspose.cells.FontUnderlineType.HEAVY:
                case com.aspose.cells.FontUnderlineType.SINGLE:
                case com.aspose.cells.FontUnderlineType.WAVE:
                case com.aspose.cells.FontUnderlineType.WAVY_DOUBLE:
                case com.aspose.cells.FontUnderlineType.WAVY_HEAVY:
                case com.aspose.cells.FontUnderlineType.WORDS:
                    c.addClass("u").setUnderline(true);
                    break;
                default:
            }

            switch (a.getStyle().getFont().getStrikeType()) {
                case com.aspose.cells.TextStrikeType.SINGLE:
                case com.aspose.cells.TextStrikeType.DOUBLE:
                    c.addClass("sts");
                    break;
                default:
            }

            switch (a.getStyle().getFont().getCapsType()) {
                case com.aspose.cells.TextCapsType.ALL:
                    c.addClass("uc");
                    break;
                case com.aspose.cells.TextCapsType.SMALL:
                    c.addClass("sc");
                    break;
                default:
            }

            style
                    .append("font-size:")
                    .append(a.getStyle().getFont().getSize())
                    .append("pt;");

            switch (a.getStyle().getHorizontalAlignment()) {
                case com.aspose.cells.TextAlignmentType.GENERAL:
                case com.aspose.cells.TextAlignmentType.LEFT:
                    c.addClass("al");
                    break;
                case com.aspose.cells.TextAlignmentType.RIGHT:
                    c.addClass("ar");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER_ACROSS:
                case com.aspose.cells.TextAlignmentType.CENTER:
                    c.addClass("ac");
                    break;
                case com.aspose.cells.TextAlignmentType.JUSTIFY:
                    c.addClass("aj");
                    break;
                default:
            }

            switch (a.getStyle().getVerticalAlignment()) {
                case com.aspose.cells.TextAlignmentType.TOP:
                    c.addClass("at");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER:
                    c.addClass("am");
                    break;
                case com.aspose.cells.TextAlignmentType.BOTTOM:
                    c.addClass("ab");
                    break;
                default:
            }

            com.aspose.cells.Color cellTextColor = a.getStyle().getFont().getColor();
            style.append("color:")
                    .append(asposeColorToCssColor(cellTextColor, true))
                    .append(";");

        } catch (Exception x) {
            LOGGER.throwing(null, null, x);
            msg.sendMessage("Cell style error", x.getMessage());
        }

        c.setStyle(style.toString());

        return c;
    }

    private String asposeColorToCssColor(com.aspose.cells.Color color, boolean emptyIsBlack) {
        Color c = asposeColorToAwtColor(color, emptyIsBlack);

        return new StringBuffer("rgb(")
                .append(c.getRed())
                .append(",")
                .append(c.getGreen())
                .append(",")
                .append(c.getBlue())
                .append(")").toString();
    }

    private Color asposeColorToAwtColor(com.aspose.cells.Color color, boolean emptyIsBlack) {
        if (color.isEmpty()) {
            if (emptyIsBlack) {
                return Color.BLACK;
            } else {
                return Color.WHITE;
            }
        }

        return new Color(color.toArgb());
    }
}
