package com.aspose.spreadsheeteditor;

import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.UUID;
import java.util.logging.Logger;
import javax.enterprise.context.ApplicationScoped;
import javax.inject.Inject;
import javax.inject.Named;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "loader")
@ApplicationScoped
public class LoaderService {

    static {
        AsposeLicense.load();
    }

    private static final Logger LOGGER = Logger.getLogger(LoaderService.class.getName());

    private HashMap<String, com.aspose.cells.Workbook> workbooks = new HashMap<>(); // NOSONAR

    @Inject
    private MessageService msg;

    @Inject
    private CellsService cells;

    public String fromBlank() {
        com.aspose.cells.Workbook w = new com.aspose.cells.Workbook();

        String key = generateKey();
        workbooks.put(key, w);

        buildCellsCache(key);
        buildColumnWidthCache(key);
        buildRowHeightCache(key);

        return key;
    }

    public String fromUrl(String url) {
        com.aspose.cells.Workbook w;

        try (InputStream i = new URL(url).openStream()) {
            w = new com.aspose.cells.Workbook(i);
        } catch (MalformedURLException murlx) {
            LOGGER.throwing(null, null, murlx);
            msg.sendMessage("The specified URL is invalid.", url);
            return null;
        } catch (IOException iox) {
            LOGGER.throwing(null, null, iox);
            msg.sendMessage("Sorry, there was a problem opening the specified file.", iox.getMessage());
            return null;
        } catch (Exception x) {
            LOGGER.throwing(null, null, x);
            msg.sendMessage("Something went wrong", x.getMessage());
            return null;
        }

        String key = generateKey();
        workbooks.put(key, w);

        buildCellsCache(key);
        buildColumnWidthCache(key);
        buildRowHeightCache(key);

        return key;
    }

    public String fromInputStream(InputStream s, String name) {
        com.aspose.cells.Workbook w;

        try (InputStream i = s) {
            w = new com.aspose.cells.Workbook(i);
        } catch (IOException iox) {
            LOGGER.throwing(null, null, iox);
            msg.sendMessage("Could not read the file from source", name);
            return null;
        } catch (Exception x) {
            LOGGER.throwing(null, null, x);
            msg.sendMessage("Could not load the workbook", name);
            return null;
        }

        String key = generateKey();
        workbooks.put(key, w);

        buildCellsCache(key);
        buildColumnWidthCache(key);
        buildRowHeightCache(key);

        return key;
    }

    public com.aspose.cells.Workbook get(String id) {
        return workbooks.get(id);
    }

    public void unload(String id) {
        workbooks.remove(id);
    }

    public void buildCellsCache(String key) {
        com.aspose.cells.Workbook wb = workbooks.get(key);
        com.aspose.cells.Worksheet ws = wb.getWorksheets().get(wb.getWorksheets().getActiveSheetIndex());
        int maxColumn = ws.getCells().getMaxColumn() + 1;
        maxColumn = maxColumn + 26 - (maxColumn % 26);
        int maxRow = 20 + ws.getCells().getMaxRow() + 1;
        maxRow = maxRow + 10 - (maxRow % 10);

        ArrayList<Column> columns = new ArrayList<>(maxColumn); // NOSONAR
        ArrayList<Row> rows = new ArrayList<>(maxRow); // NOSONAR

        for (int i = 0; i < maxColumn; i++) {
            columns.add(i, new Column(i, com.aspose.cells.CellsHelper.columnIndexToName(i)));
        }

        for (int i = 0; i < maxRow; i++) {
            rows.add(i, new Row.Builder().setId(i).build());
        }

        for (Object o : ws.getCells()) {
            com.aspose.cells.Cell c = (com.aspose.cells.Cell) o;
            rows.get(c.getRow()).putCell(c.getColumn(), cells.fromAsposeCell(c));
        }

        for (int i = 0; i < maxRow; i++) {
            for (int j = 0; j < maxColumn; j++) {
                String col = com.aspose.cells.CellsHelper.columnIndexToName(j);
                if (!rows.get(i).getCellsMap().containsKey(col)) {
                    rows.get(i).putCell(col, cells.fromBlank(j, i));
                }
            }
        }

        cells.putColumns(key, columns);
        cells.putRows(key, rows);
    }

    public void buildColumnWidthCache(String key) {
        com.aspose.cells.Workbook wb = workbooks.get(key);
        com.aspose.cells.Worksheet ws = wb.getWorksheets().get(wb.getWorksheets().getActiveSheetIndex());

        ArrayList<Integer> columnWidth = new ArrayList<>(); // NOSONAR
        for (int i = 0; i < cells.getColumns(key).size(); i++) {
            columnWidth.add(i, ws.getCells().getColumnWidthPixel(i));
        }
        cells.putColumnWidth(key, columnWidth);
    }

    public void buildRowHeightCache(String key) {
        com.aspose.cells.Workbook wb = workbooks.get(key);
        com.aspose.cells.Worksheet ws = wb.getWorksheets().get(wb.getWorksheets().getActiveSheetIndex());

        ArrayList<Integer> rowHeight = new ArrayList<>(); // NOSONAR

        for (int i = 0; i < cells.getRows(key).size(); i++) {
            rowHeight.add(i, ws.getCells().getRowHeightPixel(i));
        }

        cells.putRowHeight(key, rowHeight);
    }

    private String generateKey() {
        return UUID.randomUUID().toString();
    }
}
