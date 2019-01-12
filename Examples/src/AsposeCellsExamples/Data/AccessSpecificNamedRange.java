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

		// Getting the specified named range
		Range namedRange = worksheets.getRangeByName("TestRange");

		// Print message
		System.out.println("Named Range : " + namedRange.getRefersTo());

	}
}
