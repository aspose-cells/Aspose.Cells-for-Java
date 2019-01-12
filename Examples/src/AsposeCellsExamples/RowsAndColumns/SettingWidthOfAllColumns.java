package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class SettingWidthOfAllColumns {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingWidthOfAllColumns.class) + "RowsAndColumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Setting the width of all columns in the worksheet to 20.5
		worksheet.getCells().setStandardWidth(20.5f);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "SettingWidthOfAllColumns_out.xls");
	}
}
