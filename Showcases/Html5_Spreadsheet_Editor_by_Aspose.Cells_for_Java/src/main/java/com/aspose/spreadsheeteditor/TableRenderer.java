package com.aspose.spreadsheeteditor;

import java.io.IOException;
import javax.faces.context.FacesContext;
import javax.faces.context.ResponseWriter;
import org.primefaces.component.api.UIColumn;
import org.primefaces.component.datatable.DataTable;
import org.primefaces.component.datatable.DataTableRenderer;

/**
 *
 * @author Saqib Masood
 */
public class TableRenderer
        extends DataTableRenderer {

    @Override
    protected void encodeCell(FacesContext context, DataTable table, UIColumn column, String clientId, boolean selected)
            throws IOException {

        if (!column.isRendered()) {
            return;
        }

        ResponseWriter writer = context.getResponseWriter();
        boolean selectionEnabled = column.getSelectionMode() != null;
        String style = column.getStyle();
        String styleClass = selectionEnabled ? DataTable.SELECTION_COLUMN_CLASS : (column.getCellEditor() != null) ? DataTable.EDITABLE_COLUMN_CLASS : null;
        String userStyleClass = column.getStyleClass();
        styleClass = userStyleClass == null ? styleClass : (styleClass == null) ? userStyleClass : styleClass + " " + userStyleClass;

        writer.startElement("td", null);
        writer.writeAttribute("role", "gridcell", null);
        if (style != null) {
            writer.writeAttribute("style", style, null);
        }
        if (styleClass != null) {
            writer.writeAttribute("class", styleClass, null);
        }

        if (column.getColspan() > 1) {
            writer.writeAttribute("colspan", column.getColspan(), null);
        }

        if (column.getRowspan() > 1) {
            writer.writeAttribute("rowspan", column.getRowspan(), null);
        }

        if (selectionEnabled) {
            encodeColumnSelection(context, table, clientId, column, selected);
        }

        column.encodeAll(context);

        writer.endElement("td");
    }

}

