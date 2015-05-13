package com.aspose.spreadsheeteditor;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;
import javax.annotation.PostConstruct;
import javax.annotation.PreDestroy;
import javax.faces.context.FacesContext;
import javax.faces.view.ViewScoped;
import javax.inject.Inject;
import javax.inject.Named;
import org.primefaces.event.FileUploadEvent;
import org.primefaces.model.DefaultStreamedContent;
import org.primefaces.model.StreamedContent;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "workbook")
@ViewScoped
public class WorkbookService implements Serializable {

    private static final Logger LOGGER = Logger.getLogger(WorkbookService.class.getName());

    private String current;
    private Format outputFormat = Format.XLSX;
    private String sourceUrl;

    @Inject
    private MessageService msg;

    @Inject
    private LoaderService loader;

    @Inject
    private CellsService cells;

    @PostConstruct
    /*private*/ void init() {
        String requestedSourceUrl = FacesContext.getCurrentInstance().getExternalContext().getRequestParameterMap().get("url");
        if (requestedSourceUrl != null) {
            try {
                this.sourceUrl = new URL(requestedSourceUrl).toString();
                this.loadFromUrl();
            } catch (MalformedURLException x) {
                LOGGER.throwing(null, null, x);
                msg.sendMessageDialog("The specified URL is invalid", requestedSourceUrl);
            }
        }
    }

    @PreDestroy
    /*private*/ void destroy() {
        loader.unload(this.current);
        cells.removeColumnWidth(this.current);
        cells.removeRowHeight(this.current);
        cells.removeColumns(this.current);
        cells.removeRows(this.current);
    }

    public boolean isLoaded() {
        return this.current != null;
    }

    public Format getOutputFormat() {
        return this.outputFormat;
    }

    public void setOutputFormat(Format outputFormat) {
        this.outputFormat = outputFormat;
    }

    public void loadBlank() {
        this.current = loader.fromBlank();
    }

    public StreamedContent getOutputFile(int saveFormat) {
        if (!isLoaded()) {
            return null;
        }

        byte[] buf;
        String ext = getExtensionForSaveFormat(saveFormat);

        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            getAsposeWorkbook().save(out, saveFormat);
            buf = out.toByteArray();
        } catch (Exception x) {
            LOGGER.throwing(null, null, x);
            msg.sendMessageDialog("Could not export", x.getMessage());
            return null;
        }

        return new DefaultStreamedContent(new ByteArrayInputStream(buf), "application/octet-stream", "Spreadsheet." + ext);
    }

    public void loadFromUrl() {
        if (this.sourceUrl == null) {
            return;
        }

        this.current = loader.fromUrl(this.sourceUrl);
    }

    public void onFileUpload(FileUploadEvent e) {
        try {
            this.current = loader.fromInputStream(e.getFile().getInputstream(), e.getFile().getFileName());
        } catch (IOException x) {
            LOGGER.throwing(null, null, x);
            msg.sendMessage("Could not read the uploaded file", x.getMessage());
        }
    }

    public com.aspose.cells.Workbook getAsposeWorkbook() {
        return loader.get(this.current);
    }

    public com.aspose.cells.WorksheetCollection getAsposeWorksheets() {
        return getAsposeWorkbook().getWorksheets();
    }

    /**
     * Get ID of loaded workbook
     *
     * @return A workbook ID which is acceptable by <code>LoaderService</code>
     */
    public String getCurrent() {
        return current;
    }

    public void setCurrent(String id) {
        this.current = id;
    }

    public String getSourceUrl() {
        return sourceUrl;
    }

    public void setSourceUrl(String sourceUrl) {
        this.sourceUrl = sourceUrl;
    }

    public List<String> getSheets() {
        List<String> list = new ArrayList<>();
        if (this.isLoaded()) {
            for (int i = 0; i < getAsposeWorksheets().getCount(); i++) {
                list.add(String.valueOf(getAsposeWorksheets().get(i).getName()));
            }
        }
        return list;
    }

    public String getActiveSheet() {
        if (this.isLoaded()) {
            int i = getAsposeWorksheets().getActiveSheetIndex();
            return getAsposeWorksheets().get(i).getName();
        }
        return null;
    }

    /**
     * Switch to `name`d sheet. If there does not exist any sheet with the given
     * name, the active sheet is renamed to the given name.
     *
     * The rule is derived from the following use-case.
     *
     * If the user select a sheet from drop-down menu, this means that the sheet
     * already exist. So we can switch to that sheet. If the sheet does not
     * exist, we can say that the user has not selected it from drop-down but
     * directly modified the name of existing in the input text box.
     *
     * @param name Worksheet name
     */
    public void setActiveSheet(String name) {
        com.aspose.cells.Worksheet ws = getAsposeWorksheets().get(name);
        if (ws != null) {
            int i = ws.getIndex();
            getAsposeWorksheets().setActiveSheetIndex(i);
        } else {
            com.aspose.cells.Workbook wb = getAsposeWorkbook();
            wb.getWorksheets().get(wb.getWorksheets().getActiveSheetIndex()).setName(name);
        }

        purge();
    }

    public void onAddNewSheet() {
        if (isLoaded()) {
            try {
                int i = getAsposeWorksheets().add();
                getAsposeWorksheets().setActiveSheetIndex(i);
                purge();
            } catch (com.aspose.cells.CellsException cx) {
                LOGGER.throwing(null, null, cx);
                msg.sendMessage("New Worksheet", cx.getMessage());
            }
        }
    }

    public void onRemoveActiveSheet() {
        if (isLoaded()) {
            try {
                int i = getAsposeWorksheets().getActiveSheetIndex();
                getAsposeWorksheets().removeAt(i);
                if (getAsposeWorksheets().getCount() == 0) {
                    int j = getAsposeWorksheets().add();
                    getAsposeWorksheets().setActiveSheetIndex(j);
                }
                purge();
            } catch (com.aspose.cells.CellsException cx) {
                LOGGER.throwing(null, null, cx);
                msg.sendMessage("Could not remove sheet", cx.getMessage());
            }
        }
    }

    public void purge() {
        loader.buildCellsCache(this.current);
    }

    private String getExtensionForSaveFormat(int saveFormat) {
        String ext = null;

        switch (saveFormat) {
            case com.aspose.cells.SaveFormat.EXCEL_97_TO_2003:
                ext = "xls";
                break;
            case com.aspose.cells.SaveFormat.XLSX:
                ext = "xlsx";
                break;
            case com.aspose.cells.SaveFormat.XLSM:
                ext = "xlsm";
                break;
            case com.aspose.cells.SaveFormat.XLSB:
                ext = "xlsb";
                break;
            case com.aspose.cells.SaveFormat.XLTX:
                ext = "xltx";
                break;
            case com.aspose.cells.SaveFormat.XLTM:
                ext = "xltm";
                break;
            case com.aspose.cells.SaveFormat.SPREADSHEET_ML:
                ext = "xml";
                break;
            case com.aspose.cells.SaveFormat.PDF:
                ext = "pdf";
                break;
            case com.aspose.cells.SaveFormat.ODS:
                ext = "ods";
                break;
            default:
        }

        return ext;
    }
}
