package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.PivotField;
import com.aspose.cells.PivotFieldCollection;
import com.aspose.cells.PivotFieldDataDisplayFormat;
import com.aspose.cells.PivotFieldSubtotalType;
import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotItemPosition;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableAutoFormatType;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.PivotTableStyleType;
import com.aspose.cells.PrintOrderType;
import com.aspose.cells.Style;
import com.aspose.cells.TableStyle;
import com.aspose.cells.TableStyleElement;
import com.aspose.cells.TableStyleElementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetPivotTableFormatOptions {

	public static void main(String[] args) throws Exception {
		// The path to the resource directory
		String dataDir = Utils.getSharedDataDir(SetPivotTableFormatOptions.class) + "PivotTable/";

		Workbook workbook = new Workbook(dataDir + "PivotTable.xls");
		Worksheet sheet = workbook.getWorksheets().get(0);
		PivotTableCollection pivotTables = sheet.getPivotTables();
		PivotTable pivotTable = pivotTables.get(0);

		// Set the AutoFormat and PivotTableStyle Types
		setAutoFormatAndPivotTableStyleTypes(pivotTable);

		// Set Format Options
		setFormatOptions(pivotTable);

		// Set Row, Column and Page Fields Format
		setRowColumnAndPageFieldsFormat(pivotTable);
		
		// Modify a Pivot Table's Quick Style
		modifyAPivotTableQuickStyle(workbook, dataDir);
		
		// Clearing Pivot Fields
		clearPivotFields(workbook, dataDir);
	}

	public static void setAutoFormatAndPivotTableStyleTypes(PivotTable pivotTable) {

		//PivotTable report is automatically formatted for Excel 2003 formats
		pivotTable.setAutoFormat(true);

		//Setting the PivotTable autoformat type.
		pivotTable.setAutoFormatType(PivotTableAutoFormatType.CLASSIC);

		//Setting the PivotTable's Styles for Excel 2007/2010 formats e.g. XLSX.
		pivotTable.setPivotTableStyleType(PivotTableStyleType.PIVOT_TABLE_STYLE_LIGHT_1);
	}

	public static void setFormatOptions(PivotTable pivotTable) {

		//Dragging the third field to the data area.
		pivotTable.addFieldToArea(PivotFieldType.DATA, 2);

		//Show grand totals for rows.
		pivotTable.setRowGrand(true);

		//Show grand totals for columns.
		pivotTable.setColumnGrand(true);

		//Display a custom string in cells that contain null values.
		pivotTable.setDisplayNullString(true);
		pivotTable.setNullString("null");

		//Setting the layout
		pivotTable.setPageFieldOrder(PrintOrderType.DOWN_THEN_OVER);
	}

	public static void setRowColumnAndPageFieldsFormat(PivotTable pivotTable) {
		//Accessing the row fields.
		PivotFieldCollection pivotFields = pivotTable.getRowFields();

		//Accessing the first row field in the row fields.
		PivotField pivotField = pivotFields.get(0);

		//Setting Subtotals.
		pivotField.setSubtotals(PivotFieldSubtotalType.SUM, true);
		pivotField.setSubtotals(PivotFieldSubtotalType.COUNT, true);

		//Setting autosort options.
		//Setting the field auto sort.
		pivotField.setAutoSort(true);

		//Setting the field auto sort ascend.
		pivotField.setAscendSort(true);

		//Setting the field auto sort using the field itself.
		pivotField.setAutoSortField(-1);

		//Setting autoShow options.
		//Setting the field auto show.
		pivotField.setAutoShow(true);

		//Setting the field auto show ascend.
		pivotField.setAscendShow(false);

		//Setting the auto show using field(data field).
		pivotField.setAutoShowField(0);
	}

	@SuppressWarnings("deprecation")
	public static void setDataFieldsFormat(PivotTable pivotTable) {
		//Accessing the data fields.
		PivotFieldCollection pivotFields = pivotTable.getDataFields();

		//Accessing the first data field in the data fields.
		PivotField pivotField = pivotFields.get(0);

		//Setting data display format
		pivotField.setDataDisplayFormat(PivotFieldDataDisplayFormat.PERCENTAGE_OF);

		//Setting the base field.
		pivotField.setBaseField(1);

		//Setting the base item.
		pivotField.setBaseItem(PivotItemPosition.NEXT);

		//Setting number format
		pivotField.setNumber(10);
	}
	
	public static void modifyAPivotTableQuickStyle(Workbook workbook, String dataDir) throws Exception {
		
		//Add Pivot Table style
		Style style1 = workbook.createStyle();
		Font font1 = style1.getFont();
		font1.setColor(Color.getRed());
		
		Style style2 = workbook.createStyle();
		Font font2 = style2.getFont();
		font2.setColor(Color.getBlue());
		
		int i = workbook.getWorksheets().getTableStyles().addPivotTableStyle("tt");
		
		//Get and Set the table style for different categories
		TableStyle ts = workbook.getWorksheets().getTableStyles().get(i);
		
		int index = ts.getTableStyleElements().add(TableStyleElementType.FIRST_COLUMN);
		TableStyleElement e = ts.getTableStyleElements().get(index);
		e.setElementStyle(style1);
		
		index = ts.getTableStyleElements().add(TableStyleElementType.GRAND_TOTAL_ROW);
		e = ts.getTableStyleElements().get(index);
		e.setElementStyle(style2);

		//Set Pivot Table style name
		PivotTable pt = workbook.getWorksheets().get(0).getPivotTables().get(0);
		pt.setPivotTableStyleName ("tt");

		//Save the file.
		workbook.save(dataDir + "PivotTableQuickStyle_out.xlsx");
	}
	
	public static void clearPivotFields(Workbook workbook, String dataDir) throws Exception {
		
		//Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		//Get the pivot tables in the sheet
		PivotTableCollection pivotTables = sheet.getPivotTables();

		//Get the first PivotTable
		PivotTable pivotTable = pivotTables.get(0);

		//Clear all the data fields
		pivotTable.getDataFields().clear();

		//Add new data field
		pivotTable.addFieldToArea(PivotFieldType.DATA, "Betrag Netto FW");

		//Set the refresh data flag on
		pivotTable.setRefreshDataFlag(false);

		//Refresh and calculate the pivot table data
		pivotTable.refreshData();
		pivotTable.calculateData();

		//Save the Excel file
		workbook.save(dataDir + "ClearPivotFields_out.xlsx");
	}
}
