package AsposeCellsExamples.PivotTables;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class PivotTableSortAndHide {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the output directory.
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "PivotTableHideAndSortSample.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		PivotTable pivotTable = worksheet.getPivotTables().get(0);
		CellArea dataBodyRange = pivotTable.getDataBodyRange();
		int currentRow = 3;
		int rowsUsed = dataBodyRange.EndRow;

		// Sorting score in descending
		PivotField field = pivotTable.getRowFields().get(0);
		field.setAutoSort(true);
		field.setAscendSort(false);
		field.setAutoSortField(0);

		pivotTable.refreshData();
		pivotTable.calculateData();

		// Hiding rows with score less than 60
		while (currentRow < rowsUsed)
		{
			Cell cell = worksheet.getCells().get(currentRow, 1);
			double score = (double) cell.getValue();
			if (score < 60)
			{
				worksheet.getCells().hideRow(currentRow);
			}
			currentRow++;
		}

		pivotTable.refreshData();
		pivotTable.calculateData();

		// Saving the Excel file
		workbook.save(outputDir + "PivotTableHideAndSort_out.xlsx");
		// ExEnd:1

		System.out.println("PivotTableSortAndHide executed successfully.");
	}
}
