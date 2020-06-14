package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class WrapTextinCell {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WrapTextinCell.class) + "TechnicalArticles/";

		// Create Workbook Object
		Workbook wb = new Workbook();

		// Open first Worksheet in the workbook
		Worksheet ws = wb.getWorksheets().get(0);

		// Get Worksheet Cells Collection
		Cells cell = ws.getCells();

		// Increase the width of First Column Width
		cell.setColumnWidth(0, 35);

		// Increase the height of first row
		cell.setRowHeight(0, 65);

		// Add Text to the First Cell
		cell.get(0, 0).setValue("I am using the latest version of Aspose.Cells to test this functionality");

		// Get Cell's Style
		Style style = cell.get(0, 0).getStyle();

		// Set Text Wrap property to true
		style.setTextWrapped(true);

		// Set Cell's Style
		cell.get(0, 0).setStyle(style);

		// Save Excel File
		wb.save(dataDir + "WrapTextinCell_out.xls");
		// ExEnd:1
	}
}
