package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class LockCell {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(LockCell.class) + "Worksheets/";

		// Instantiating a Workbook object by excel file path
		Workbook excel = new Workbook(dataDir + "Book1.xlsx");

		WorksheetCollection worksheets = excel.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		worksheet.getCells().get("A1").getStyle().setLocked(true);

		// Saving the modified Excel file Excel XP format
		excel.save(dataDir + "LockCell_out.xls");

		// Print Message
		System.out.println("Cell Locked successfully.");

	}
}
