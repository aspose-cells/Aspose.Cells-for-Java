package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class SettingHeightAllRows {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingHeightAllRows.class) + "RowsAndColumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Setting the height of all rows in the worksheet to 15
		worksheet.getCells().setStandardHeight(15f);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "SettingHeightAllRows_out.xls");
	}
}
