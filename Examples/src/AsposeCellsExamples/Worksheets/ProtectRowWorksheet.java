package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ProtectRowWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ProtectRowWorksheet.class) + "Worksheets/";

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
			style = sheet.getCells().getRows().get(i).getStyle();
			style.setLocked(false);
			flag = new StyleFlag();
			flag.setLocked(true);
			sheet.getCells().getRows().get(i).applyStyle(style, flag);
		}

		// Get the first Roww style.
		style = sheet.getCells().getRows().get(1).getStyle();

		// Lock it.
		style.setLocked(true);

		// Instantiate the flag.
		flag = new StyleFlag();

		// Set the lock setting.
		flag.setLocked(true);

		// Apply the style to the first row.
		sheet.getCells().getRows().get(1).applyStyle(style, flag);
		sheet.protect(ProtectionType.ALL);

		// Save the excel file.
		wb.save(dataDir + "ProtectRowWorksheet_out.xls");

		// Print Message
		System.out.println("Row protected successfully.");

	}
}
