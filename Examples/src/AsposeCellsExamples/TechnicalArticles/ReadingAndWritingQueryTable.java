package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.QueryTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ReadingAndWritingQueryTable {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ReadingAndWritingQueryTable.class) + "TechnicalArticles/";
		// Create workbook from source excel file
		Workbook workbook = new Workbook(dataDir + "SampleQT.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first Query Table
		QueryTable queryTable = worksheet.getQueryTables().get(0);

		// Print Query Table Data
		System.out.println("Adjust Column Width: " + queryTable.getAdjustColumnWidth());
		System.out.println("Preserve Formatting: " + queryTable.getPreserveFormatting());

		// Now set Preserve Formatting to true
		queryTable.setPreserveFormatting(true);

		// Save the workbook
		workbook.save(dataDir + "RAWQueryTable_out.xlsx");


	}
}
