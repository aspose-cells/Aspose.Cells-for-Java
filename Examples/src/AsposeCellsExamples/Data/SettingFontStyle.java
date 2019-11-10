package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SettingFontStyle {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the output directory.
		String outputDir = Utils.Get_OutputDirectory();

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a new worksheet to the Excel object
		int i = workbook.getWorksheets().add();

		// Obtaining the reference of the newly added worksheet by passing its sheet index
		Worksheet worksheet = workbook.getWorksheets().get(i);

		// Accessing the "A1" cell from the worksheet
		Cell cell = worksheet.getCells().get("A1");

		// Adding some value to the "A1" cell
		cell.putValue("Hello Aspose!");

		// Obtaining the style of the cell
		Style style = cell.getStyle();
		// Setting the font weight to bold
		style.getFont().setBold(true);
		// Applying the style to the cell
		cell.setStyle(style);

		// Saving the Excel file
		workbook.save(outputDir + "book1.out.xlsx", SaveFormat.XLSX);
		// ExEnd:1
	}
}
