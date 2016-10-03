package com.aspose.cells.examples.data;

import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingDisplayFormat {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingDisplayFormat.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the first (default) worksheet by passing its sheet index
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Adding a new Style to the styles collection of the Workbook object
		Style style = workbook.createStyle();

		// Setting the Number property to 4 which corresponds to the pattern #,##0.00
		style.setNumber(4);

		// Creating an object of StyleFlag
		StyleFlag flag = new StyleFlag();

		// Setting NumberFormat property to true so that only this aspect takes effect from Style object
		flag.setNumberFormat(true);

		// Applying style to the first row of the worksheet
		worksheet.getCells().getRows().get(0).applyStyle(style, flag);

		// Re-initializing the Style object
		style = workbook.createStyle();

		// Setting the Custom property to the pattern d-mmm-yy
		style.setCustom("d-mmm-yy");

		// Applying style to the first column of the worksheet
		worksheet.getCells().getColumns().get(0).applyStyle(style, flag);

		// Saving spreadsheet on disc
		workbook.save(dataDir + "SDisplayFormat_out.xlsx");
	}
}
