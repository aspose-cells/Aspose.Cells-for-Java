package AsposeCellsExamples.PivotTables;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class PivotTableSaveInODS {
	public static void main(String[] args) throws Exception {
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		// ExStart:1
		String outputDir = Utils.Get_OutputDirectory();

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the newly added worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		Cells cells = sheet.getCells();

		// Setting the value to the cells
		Cell cell = cells.get("A1");
		cell.putValue("Sport");
		cell = cells.get("B1");
		cell.putValue("Quarter");
		cell = cells.get("C1");
		cell.putValue("Sales");

		cell = cells.get("A2");
		cell.putValue("Golf");
		cell = cells.get("A3");
		cell.putValue("Golf");
		cell = cells.get("A4");
		cell.putValue("Tennis");
		cell = cells.get("A5");
		cell.putValue("Tennis");
		cell = cells.get("A6");
		cell.putValue("Tennis");
		cell = cells.get("A7");
		cell.putValue("Tennis");
		cell = cells.get("A8");
		cell.putValue("Golf");

		cell = cells.get("B2");
		cell.putValue("Qtr3");
		cell = cells.get("B3");
		cell.putValue("Qtr4");
		cell = cells.get("B4");
		cell.putValue("Qtr3");
		cell = cells.get("B5");
		cell.putValue("Qtr4");
		cell = cells.get("B6");
		cell.putValue("Qtr3");
		cell = cells.get("B7");
		cell.putValue("Qtr4");
		cell = cells.get("B8");
		cell.putValue("Qtr3");

		cell = cells.get("C2");
		cell.putValue(1500);
		cell = cells.get("C3");
		cell.putValue(2000);
		cell = cells.get("C4");
		cell.putValue(600);
		cell = cells.get("C5");
		cell.putValue(1500);
		cell = cells.get("C6");
		cell.putValue(4070);
		cell = cells.get("C7");
		cell.putValue(5000);
		cell = cells.get("C8");
		cell.putValue(6430);

		PivotTableCollection pivotTables = sheet.getPivotTables();

		// Adding a PivotTable to the worksheet
		int index = pivotTables.add("=A1:C8", "E3", "PivotTable2");

		// Accessing the instance of the newly added PivotTable
		PivotTable pivotTable = pivotTables.get(index);

		// Unshowing grand totals for rows.
		pivotTable.setRowGrand(false);

		// Draging the first field to the row area.
		pivotTable.addFieldToArea(PivotFieldType.ROW, 0);

		// Draging the second field to the column area.
		pivotTable.addFieldToArea(PivotFieldType.COLUMN, 1);

		// Draging the third field to the data area.
		pivotTable.addFieldToArea(PivotFieldType.DATA, 2);

		pivotTable.calculateData();

		// Saving the ODS file
		workbook.save(outputDir + "PivotTableSaveInODS_out.ods");
		// ExEnd:1

		System.out.println("PivotTableSaveInODS executed successfully.");
	}
}
