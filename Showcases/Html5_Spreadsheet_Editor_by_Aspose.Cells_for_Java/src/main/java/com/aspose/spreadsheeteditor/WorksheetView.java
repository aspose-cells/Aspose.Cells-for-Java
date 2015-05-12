package com.aspose.spreadsheeteditor;

import java.awt.Color;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;
import javax.faces.context.FacesContext;
import javax.faces.view.ViewScoped;
import javax.inject.Inject;
import javax.inject.Named;
import org.primefaces.context.RequestContext;
import org.primefaces.event.CellEditEvent;
import org.primefaces.event.ColumnResizeEvent;

@Named(value = "worksheet")
@ViewScoped
public class WorksheetView implements Serializable {

    private static final Logger LOGGER = Logger.getLogger(WorksheetView.class.getName());

    private int currentColumnId;
    private int currentRowId;
    private String currentCellClientId;

    @Inject
    private WorkbookService workbook;

    @Inject
    private CellsService cells;

    @Inject
    private MessageService msg;

    @Inject
    private LoaderService loader;

    @Inject
    FormattingService formatting;

    public boolean isLoaded() {
        return workbook.isLoaded();
    }

    public int getDefaultColumnWidth() {
        try {
            return getAsposeWorksheet().getCells().getStandardWidthPixels();
        } catch (com.aspose.cells.CellsException | NullPointerException x) {
            LOGGER.throwing(null, null, x);
            return 64;
        }
    }

    public List<Integer> getColumnWidth() {
        return cells.getColumnWidth(workbook.getCurrent());
    }

    public List<Integer> getRowHeight() {
        return cells.getRowHeight(workbook.getCurrent());
    }

    private com.aspose.cells.Worksheet getAsposeWorksheet() {
        com.aspose.cells.Workbook w = workbook.getAsposeWorkbook();
        return w.getWorksheets().get(w.getWorksheets().getActiveSheetIndex());
    }

    public List<Column> getColumns() {
        return cells.getColumns(workbook.getCurrent());
    }

    public List<Row> getRows() {
        return cells.getRows(workbook.getCurrent());
    }

    public void applyCellFormatting() {
        if (!isLoaded()) {
            return;
        }

        com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(currentRowId, currentColumnId);
        com.aspose.cells.Style s = c.getStyle();

        s.getFont().setBold(formatting.isBoldOptionEnabled());
        s.getFont().setItalic(formatting.isItalicOptionEnabled());
        s.getFont().setUnderline(formatting.isUnderlineOptionEnabled() ? com.aspose.cells.FontUnderlineType.SINGLE : com.aspose.cells.FontUnderlineType.NONE);
        s.getFont().setName(formatting.getFontSelectionOption());
        s.getFont().setSize(formatting.getFontSizeOption());
        switch (formatting.getAlignSelectionOption()) {
            case "al":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.LEFT);
                break;
            case "ac":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.CENTER);
                break;
            case "ar":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.RIGHT);
                break;
            case "aj":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.JUSTIFY);
                break;
            default:
        }
        try {
            s.getFont().setColor(com.aspose.cells.Color.fromArgb(Color.decode("0x" + formatting.getFontColorSelectionOption()).getRGB()));
        } catch (NumberFormatException x) {
            // Ignore
        }
        try {
            s.setForegroundArgbColor(Color.decode("0x" + formatting.getFillColorSelectionOption()).getRGB());
        } catch (NumberFormatException x) {
            // Ignore
        }

        c.setStyle(s);
        RequestContext.getCurrentInstance().update(currentCellClientId);
        purge();
    }

    public int getCurrentColumnId() {
        return currentColumnId;
    }

    public void setCurrentColumnId(int currentColumnId) {
        this.currentColumnId = currentColumnId;
    }

    public int getCurrentRowId() {
        return currentRowId;
    }

    public void setCurrentRowId(int currentRowId) {
        this.currentRowId = currentRowId;
    }

    public String getCurrentCellClientId() {
        return currentCellClientId;
    }

    public void setCurrentCellClientId(String currentCellClientId) {
        this.currentCellClientId = currentCellClientId;
    }

    public void onCellEdit(CellEditEvent e) {
        Cell newCell = (Cell) e.getNewValue();
        int columnId = newCell.getColumnId();
        int rowId = newCell.getRowId();

        try {
            com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(rowId, columnId);
            if (newCell.getFormula() != null) {
                c.setFormula(newCell.getFormula(), null);
            } else {
                c.putValue(newCell.getValue(), true);
            }
            cells.putCell(workbook.getCurrent(), columnId, rowId, newCell);
        } catch (com.aspose.cells.CellsException x) {
            LOGGER.throwing(null, null, x);
        }
    }

    public void onColumnResize(ColumnResizeEvent e) {
        if (!isLoaded()) {
            return;
        }

        int columnId = com.aspose.cells.CellsHelper.columnNameToIndex(e.getColumn().getHeaderText());
        try {
            getAsposeWorksheet().getCells().setColumnWidthPixel(columnId, e.getWidth());
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Could not resize column", cx.getMessage());
            return;
        }

        reloadColumnWidth(columnId);
    }

    public void addRowAbove() {
        try {
            getAsposeWorksheet().getCells().insertRows(currentRowId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Could not add row", cx.getMessage());
            return;
        }

        purge();
        reloadRowHeight(currentRowId);
    }

    public void addRowBelow() {
        if (getCurrentRowId() < 0) {
            msg.sendMessage("No cell selected", null);
            return;
        }

        int newRowId = currentRowId + 1;

        try {
            getAsposeWorksheet().getCells().insertRows(newRowId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Could not add row", cx.getMessage());
            return;
        }

        purge();
        reloadRowHeight(newRowId);
    }

    public void deleteRow() {
        try {
            getAsposeWorksheet().getCells().deleteRows(currentRowId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Could not delete row", cx.getMessage());
            return;
        }

        cells.getRows(workbook.getCurrent()).remove(currentRowId);
        getRowHeight().remove(currentRowId);
        purge();
    }

    public void addColumnBefore() {
        try {
            getAsposeWorksheet().getCells().insertColumns(getCurrentColumnId(), 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Could not add column", cx.getMessage());
            return;
        }

        reloadColumnWidth(currentColumnId);
        purge();
    }

    public void addColumnAfter() {
        int newColumnId = currentColumnId + 1;
        try {
            getAsposeWorksheet().getCells().insertColumns(newColumnId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Could not add column", cx.getMessage());
            return;
        }

        reloadColumnWidth(newColumnId);
        purge();
    }

    public void deleteColumn() {
        try {
            getAsposeWorksheet().getCells().deleteColumns(currentColumnId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            LOGGER.throwing(null, null, cx);
            msg.sendMessage("Could not delete column", cx.getMessage());
            return;
        }

        cells.getColumns(workbook.getCurrent()).remove(currentColumnId);
        getRowHeight().remove(currentColumnId);
        purge();
    }

    public void addCellShiftRight() {
        if (!isLoaded()) {
            return;
        }

        com.aspose.cells.CellArea a = new com.aspose.cells.CellArea();
        a.StartColumn = a.EndColumn = currentColumnId;
        a.StartRow = a.EndRow = currentRowId;
        getAsposeWorksheet().getCells().insertRange(a, com.aspose.cells.ShiftType.RIGHT);
        purge();
    }

    public void addCellShiftDown() {
        if (!isLoaded()) {
            return;
        }

        com.aspose.cells.CellArea a = new com.aspose.cells.CellArea();
        a.StartColumn = a.EndColumn = currentColumnId;
        a.StartRow = a.EndRow = currentRowId;
        getAsposeWorksheet().getCells().insertRange(a, com.aspose.cells.ShiftType.DOWN);
        purge();
    }

    public void removeCellShiftUp() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().deleteRange(currentRowId, currentColumnId, currentRowId, currentColumnId, com.aspose.cells.ShiftType.UP);
        purge();
    }

    public void removeCellShiftLeft() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().deleteRange(currentRowId, currentColumnId, currentRowId, currentColumnId, com.aspose.cells.ShiftType.LEFT);
        purge();
    }

    public void clearCurrentCellFormatting() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().clearFormats(currentRowId, currentColumnId, currentRowId, currentColumnId);
        reloadCell(currentColumnId, currentRowId);
        RequestContext.getCurrentInstance().update(currentCellClientId);
    }

    public void clearCurrentCellContents() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().clearContents(currentRowId, currentColumnId, currentRowId, currentColumnId);
        reloadCell(currentColumnId, currentRowId);
        RequestContext.getCurrentInstance().update(currentCellClientId);
    }

    public void clearCurrentCell() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().clearRange(currentRowId, currentColumnId, currentRowId, currentColumnId);
        reloadCell(currentColumnId, currentRowId);
        RequestContext.getCurrentInstance().update(currentCellClientId);
    }

    public int getCurrentColumnWidth() {
        return getColumnWidth().get(currentColumnId);
    }

    public void setCurrentColumnWidth(int width) {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().setColumnWidthPixel(currentColumnId, width);
        reloadColumnWidth(currentColumnId);
        RequestContext.getCurrentInstance().update("sheet");

    }

    public int getCurrentRowHeight() {
        return getRowHeight().get(currentRowId);
    }

    public void setCurrentRowHeight(int height) {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().setRowHeightPixel(currentRowId, height);
        reloadRowHeight(currentRowId);
        RequestContext.getCurrentInstance().update("sheet");
    }

    public List<String> getFonts() {
        List<String> list = new ArrayList<>();
        if (isLoaded()) {
            // TODO: get list of fonts used by the workboook
        }
        return list;
    }

    public void updatePartialView() {
        String id = FacesContext.getCurrentInstance().getExternalContext().getRequestParameterMap().get("id");
        if (id != null) {
            RequestContext.getCurrentInstance().update(id);
        }
    }

    private void reloadColumnWidth(int columnId) {
        int width = getAsposeWorksheet().getCells().getColumnWidthPixel(columnId);
        getColumnWidth().remove(columnId);
        getColumnWidth().add(columnId, width);
    }

    private void reloadRowHeight(int rowId) {
        int height = getAsposeWorksheet().getCells().getRowHeightPixel(rowId);
        getRowHeight().remove(rowId);
        getRowHeight().add(rowId, height);
    }

    private void reloadCell(int columnId, int rowId) {
        com.aspose.cells.Cell a = getAsposeWorksheet().getCells().get(rowId, columnId);
        Cell c = cells.fromAsposeCell(a);
        cells.getRows(workbook.getCurrent()).get(rowId).putCell(columnId, c);
    }

    public void purge() {
        loader.buildCellsCache(workbook.getCurrent());
    }
}
