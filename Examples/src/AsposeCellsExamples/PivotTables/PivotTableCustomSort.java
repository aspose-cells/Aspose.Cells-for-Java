package AsposeCellsExamples.PivotTables;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class PivotTableCustomSort {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the output directory.
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		Workbook wb = new Workbook(sourceDir + "SamplePivotSort.xlsx");

		// Obtaining the reference of the newly added worksheet
		Worksheet sheet = wb.getWorksheets().get(0);

		PivotTableCollection pivotTables = sheet.getPivotTables();

		// source PivotTable
		// Adding a PivotTable to the worksheet
		int index = pivotTables.add("=Sheet1!A1:C10", "E3", "PivotTable2");

		//Accessing the instance of the newly added PivotTable
		PivotTable pivotTable = pivotTables.get(index);

		// Unshowing grand totals for rows.
		pivotTable.setRowGrand(false);
		pivotTable.setColumnGrand(false);

		// Dragging the first field to the row area.
		pivotTable.addFieldToArea(PivotFieldType.ROW, 1);
		PivotField rowField = pivotTable.getRowFields().get(0);
		rowField.setAutoSort(true);
		rowField.setAscendSort(true);

		// Dragging the second field to the column area.
		pivotTable.addFieldToArea(PivotFieldType.COLUMN, 0);
		PivotField colField = pivotTable.getColumnFields().get(0);
		colField.setNumberFormat("dd/mm/yyyy");
		colField.setAutoSort(true);
		colField.setAscendSort(true);

		// Dragging the third field to the data area.
		pivotTable.addFieldToArea(PivotFieldType.DATA, 2);

		pivotTable.refreshData();
		pivotTable.calculateData();
		// end of source PivotTable


		// sort the PivotTable on "SeaFood" row field values
		// Adding a PivotTable to the worksheet
		index = pivotTables.add("=Sheet1!A1:C10", "E10", "PivotTable2");

		// Accessing the instance of the newly added PivotTable
		pivotTable = pivotTables.get(index);

		// Unshowing grand totals for rows.
		pivotTable.setRowGrand(false);
		pivotTable.setColumnGrand(false);

		// Dragging the first field to the row area.
		pivotTable.addFieldToArea(PivotFieldType.ROW, 1);
		rowField = pivotTable.getRowFields().get(0);
		rowField.setAutoSort(true);
		rowField.setAscendSort(true);

		// Dragging the second field to the column area.
		pivotTable.addFieldToArea(PivotFieldType.COLUMN, 0);
		colField = pivotTable.getColumnFields().get(0);
		colField.setNumberFormat("dd/mm/yyyy");
		colField.setAutoSort(true);
		colField.setAscendSort(true);
		colField.setAutoSortField(0);


		//Dragging the third field to the data area.
		pivotTable.addFieldToArea(PivotFieldType.DATA, 2);

		pivotTable.refreshData();
		pivotTable.calculateData();
		// end of sort the PivotTable on "SeaFood" row field values


		// sort the PivotTable on "28/07/2000" column field values
		// Adding a PivotTable to the worksheet
		index = pivotTables.add("=Sheet1!A1:C10", "E18", "PivotTable2");

		// Accessing the instance of the newly added PivotTable
		pivotTable = pivotTables.get(index);

		// Unshowing grand totals for rows.
		pivotTable.setRowGrand(false);
		pivotTable.setColumnGrand(false);
		// Dragging the first field to the row area.
		pivotTable.addFieldToArea(PivotFieldType.ROW, 1);
		rowField = pivotTable.getRowFields().get(0);
		rowField.setAutoSort(true);
		rowField.setAscendSort(true);
		rowField.setAutoSortField(0);

		// Dragging the second field to the column area.
		pivotTable.addFieldToArea(PivotFieldType.COLUMN, 0);
		colField = pivotTable.getColumnFields().get(0);
		colField.setNumberFormat("dd/mm/yyyy");
		colField.setAutoSort(true);
		colField.setAscendSort(true);


		// Dragging the third field to the data area.
		pivotTable.addFieldToArea(PivotFieldType.DATA, 2);

		pivotTable.refreshData();
		pivotTable.calculateData();
		// end of sort the PivotTable on "28/07/2000" column field values


		//Saving the Excel file
		wb.save(outputDir + "out_java.xlsx");
		PdfSaveOptions options = new PdfSaveOptions();
		options.setOnePagePerSheet(true);
		wb.save(outputDir + "out_java.pdf", options);
	}
}
