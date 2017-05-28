package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AccessSpecificNamedRange {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AccessSpecificNamedRange.class) + "Data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();

		// Accessing the first worksheet in the Excel file
		Worksheet sheet = worksheets.get(0);
		Cells cells = sheet.getCells();

		// Getting the specified named range
		Range namedRange = worksheets.getRangeByName("TestRange");

		// Print message
		System.out.println("Named Range : " + namedRange.getRefersTo());

	}
}
