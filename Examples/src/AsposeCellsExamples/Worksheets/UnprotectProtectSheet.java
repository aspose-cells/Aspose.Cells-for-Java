package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class UnprotectProtectSheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UnprotectProtectSheet.class) + "Worksheets/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		// Unprotecting the worksheet
		worksheet.unprotect("aspose");

		// Save the excel file.
		workbook.save(dataDir + "UnprotectProtectSheet_out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Worksheet unprotected successfully.");

	}
}
