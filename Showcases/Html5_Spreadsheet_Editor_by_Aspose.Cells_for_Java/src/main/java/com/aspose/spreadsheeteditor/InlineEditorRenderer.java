package com.aspose.spreadsheeteditor;

import java.io.IOException;
import javax.faces.component.UIComponent;
import javax.faces.context.FacesContext;
import javax.faces.context.ResponseWriter;
import org.primefaces.component.celleditor.CellEditor;
import org.primefaces.component.celleditor.CellEditorRenderer;
import org.primefaces.component.datatable.DataTable;

public class InlineEditorRenderer extends CellEditorRenderer {

    @Override
    public void encodeEnd(FacesContext context, UIComponent component) throws IOException {
        ResponseWriter writer = context.getResponseWriter();
        CellEditor editor = (CellEditor) component;

        writer.startElement("div", null);
        writer.writeAttribute("id", component.getClientId(context), null);
        writer.writeAttribute("class", DataTable.CELL_EDITOR_CLASS, null);

        writer.startElement("div", null);
        writer.writeAttribute("class", DataTable.CELL_EDITOR_OUTPUT_CLASS, null);
        editor.getFacet("output").encodeAll(context);
        writer.endElement("div");

        writer.startElement("div", null);
        writer.writeAttribute("class", DataTable.CELL_EDITOR_INPUT_CLASS, null);
        editor.getFacet("input").encodeAll(context);
        writer.endElement("div");

        writer.endElement("div");
    }
}
