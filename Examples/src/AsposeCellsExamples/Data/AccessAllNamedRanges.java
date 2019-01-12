package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AccessAllNamedRanges {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AccessAllNamedRanges.class) + "Data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();

		// Getting all named ranges
		Range[] namedRanges = worksheets.getNamedRanges();

		// Print message
		System.out.println("Number of Named Ranges : " + namedRanges.length);

	}
}
