package com.aspose.spreadsheeteditor;

import javax.el.ValueExpression;
import javax.faces.component.UIComponent;
import javax.faces.context.FacesContext;
import javax.faces.convert.Converter;
import javax.faces.convert.FacesConverter;
import javax.inject.Inject;

@FacesConverter("cellEditorConverter")
public class CellConverter implements Converter {

    @Inject
    private WorkbookService workbook;

    @Inject
    private CellsService cells;

    @Override
    public Object getAsObject(FacesContext context, UIComponent component, String value) {
        int columnId = (Integer) ((ValueExpression) component.getPassThroughAttributes().get("data-columnId")).getValue(context.getELContext());
        int rowId = (Integer) ((ValueExpression) component.getPassThroughAttributes().get("data-rowId")).getValue(context.getELContext());

        Cell cell = cells.getCell(workbook.getCurrent(), columnId, rowId);
        if (value.trim().startsWith("=")) {
            cell.setFormula(value.trim());
        } else {
            cell.setValue(value);
        }

        return cell;
    }

    @Override
    public String getAsString(FacesContext context, UIComponent component, Object value) {
        Cell cell = (Cell) value;

        if (cell.getFormula() != null) {
            return cell.getFormula();
        } else {
            return cell.getValue();
        }
    }
}
