package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
public class SpecifyingCustomSeparators {
	public static void main(String[] args) throws Exception {

		Workbook workbook = new Workbook();

		// Specify custom separators
		workbook.getSettings().setNumberDecimalSeparator('.');
		workbook.getSettings().setNumberGroupSeparator(' ');

		Worksheet worksheet = workbook.getWorksheets().get(0);

		Cell cell = worksheet.getCells().get("A1");
		cell.putValue(123456.789);

		Style style = cell.getStyle();
		style.setCustom("#,##0.000;[Red]#,##0.000");
		cell.setStyle(style);

		worksheet.autoFitColumns();

		workbook.save("SpecifyingCustomSeparators_out.pdf");

	}
}
