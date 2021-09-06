package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ProtectColumnWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ProtectColumnWorksheet.class) + "Worksheets/";

		// Create a new workbook.
		Workbook wb = new Workbook();

		// Create a worksheet object and obtain the first sheet.
		Worksheet sheet = wb.getWorksheets().get(0);

		// Define the style object.
		Style style;

		// Define the styleflag object.
		StyleFlag flag;

		// Loop through all the columns in the worksheet and unlock them.
		for (int i = 0; i <= 255; i++) {
			style = sheet.getCells().getColumns().get(i).getStyle();
			style.setLocked(false);
			flag = new StyleFlag();
			flag.setLocked(true);
			sheet.getCells().getColumns().get(i).applyStyle(style, flag);
		}

		// Get the first column style.
		style = sheet.getCells().getColumns().get(0).getStyle();

		// Lock it.
		style.setLocked(true);

		// Instantiate the flag.
		flag = new StyleFlag();

		// Set the lock setting.
		flag.setLocked(true);

		// Apply the style to the first column.
		sheet.getCells().getColumns().get(0).applyStyle(style, flag);
		sheet.protect(ProtectionType.ALL);

		// Save the excel file.
		wb.save(dataDir + "PColumnWorksheet_out.xls");

		// Print Message
		System.out.println("Column protected successfully.");

	}
}
