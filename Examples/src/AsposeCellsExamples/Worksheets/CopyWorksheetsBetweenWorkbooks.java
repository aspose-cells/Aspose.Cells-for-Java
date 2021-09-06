package AsposeCellsExamples.Worksheets;

import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class CopyWorksheetsBetweenWorkbooks {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(AddingPageBreaks.class) + "Worksheets/";
		// Create a Workbook.
		Workbook excelWorkbook0 = new Workbook(dataDir + "book1.xls");

		// Create another Workbook.
		Workbook excelWorkbook1 = new Workbook();

		// Copy the first sheet of the first book into second book.
		excelWorkbook1.getWorksheets().get(0).copy(excelWorkbook0.getWorksheets().get(0));

		// Save the file.
		excelWorkbook1.save(dataDir + "CWBetweenWorkbooks_out.xls");
	}
}
