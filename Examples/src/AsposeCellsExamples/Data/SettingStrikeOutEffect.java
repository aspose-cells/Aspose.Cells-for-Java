package AsposeCellsExamples.Data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class SettingStrikeOutEffect {
	public static void main(String[] args) throws Exception {
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = worksheet.getCells();

		// Adding some value to the "A1" cell
		Cell cell = cells.get("A1");
		cell.setValue("Hello Aspose!");

		// Setting the strike out effect on the font
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setStrikeout(true);
		cell.setStyle(style);
	}
}
