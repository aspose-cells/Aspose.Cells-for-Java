package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Cell;
import com.aspose.cells.CellValueFormatStrategy;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class GetCellStringValue {
	public static void main(String[] args) throws Exception {

		// Create workbook
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access cell A1
		Cell cell = worksheet.getCells().get("A1");

		// Put value inside the cell
		cell.putValue(0.012345);

		// Format the cell that it should display 0.01 instead of 0.012345
		Style style = cell.getStyle();
		style.setNumber(2);
		cell.setStyle(style);

		// Get string value as Cell Style
		String value = cell.getStringValue(CellValueFormatStrategy.CELL_STYLE);
		System.out.println(value);

		// Get string value without any formatting
		value = cell.getStringValue(CellValueFormatStrategy.NONE);
		System.out.println(value);

	}
}
