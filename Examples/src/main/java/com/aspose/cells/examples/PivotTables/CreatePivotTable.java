package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.DrawingObjects.NonPrimitiveShape;

public class CreatePivotTable {
	public static void main(String[] args) throws Exception {
		
		// The path to the resource directory
		String dataDir = Utils.getSharedDataDir(CreatePivotTable.class) + "PivotTables/";
				
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Obtaining the reference of the newly added worksheet
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet sheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = sheet.getCells();

		//Setting the value to the cells
		Cell cell = cells.get("A1");
		cell.setValue("Sport");
		cell = cells.get("B1");
		cell.setValue("Quarter");
		cell = cells.get("C1");
		cell.setValue("Sales");

		cell = cells.get("A2");
		cell.setValue("Golf");
		cell = cells.get("A3");
		cell.setValue("Golf");
		cell = cells.get("A4");
		cell.setValue("Tennis");
		cell = cells.get("A5");
		cell.setValue("Tennis");
		cell = cells.get("A6");
		cell.setValue("Tennis");
		cell = cells.get("A7");
		cell.setValue("Tennis");
		cell = cells.get("A8");
		cell.setValue("Golf");

		cell = cells.get("B2");
		cell.setValue("Qtr3");
		cell = cells.get("B3");
		cell.setValue("Qtr4");
		cell = cells.get("B4");
		cell.setValue("Qtr3");
		cell = cells.get("B5");
		cell.setValue("Qtr4");
		cell = cells.get("B6");
		cell.setValue("Qtr3");
		cell = cells.get("B7");
		cell.setValue("Qtr4");
		cell = cells.get("B8");
		cell.setValue("Qtr3");

		cell = cells.get("C2");
		cell.setValue(1500);
		cell = cells.get("C3");
		cell.setValue(2000);
		cell = cells.get("C4");
		cell.setValue(600);
		cell = cells.get("C5");
		cell.setValue(1500);
		cell = cells.get("C6");
		cell.setValue(4070);
		cell = cells.get("C7");
		cell.setValue(5000);
		cell = cells.get("C8");
		cell.setValue(6430);

		PivotTableCollection pivotTables = sheet.getPivotTables();

		//Adding a PivotTable to the worksheet
		int index = pivotTables.add("=A1:C8", "E3", "PivotTable2");

		//Accessing the instance of the newly added PivotTable
		PivotTable pivotTable = pivotTables.get(index);

		//Unshowing grand totals for rows.
		pivotTable.setRowGrand(false);

		//Dragging the first field to the row area.
		pivotTable.addFieldToArea(PivotFieldType.ROW, 0);

		//Dragging the second field to the column area.
		pivotTable.addFieldToArea(PivotFieldType.COLUMN, 1);

		//Dragging the third field to the data area.
		pivotTable.addFieldToArea(PivotFieldType.DATA, 2);

		//Saving the Excel file
		workbook.save(dataDir + "CreatePivotTable_out.xls");
	}
}
