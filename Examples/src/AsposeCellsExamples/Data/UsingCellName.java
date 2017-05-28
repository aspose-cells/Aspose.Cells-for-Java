package AsposeCellsExamples.Data;

import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class UsingCellName {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingCellName.class) + "Data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the worksheet in the Excel file
		com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(0);
		com.aspose.cells.Cells cells = worksheet.getCells();

		// Accessing a cell using its name
		com.aspose.cells.Cell cell = cells.get("A1");

		// Print message
		System.out.println("Cell Value: " + cell.getValue());

	}
}
