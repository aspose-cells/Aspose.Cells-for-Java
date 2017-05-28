package AsposeCellsExamples.Charts;

import com.aspose.cells.CellsColor;
import com.aspose.cells.Chart;
import com.aspose.cells.FillType;
import com.aspose.cells.ThemeColor;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import AsposeCellsExamples.Utils;

public class MicrosoftTheme {

	public static void main(String[] args) throws Exception {
		// ExStart:MicrosoftTheme
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(MicrosoftTheme.class) + "Charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Obtaining the reference of the first worksheet
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		Chart chart = sheet.getCharts().get(0);

		// Specify the FilFormat's type to Solid Fill of the first series
		chart.getNSeries().get(0).getArea().getFillFormat().setFillType(FillType.SOLID);

		// Get the CellsColor of SolidFill
		CellsColor cc = chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().getCellsColor();

		// Create a theme in Accent style
		cc.setThemeColor(new ThemeColor(ThemeColorType.FOLLOWED_HYPERLINK, 0.6));

		// Apply the them to the series
		chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().setCellsColor(cc);

		// Save the Excel file
		workbook.save(dataDir + "MicrosoftTheme_out.xlsx");

		// Print message
		System.out.println("MicrosoftTheme is successfully applied.");
		// ExEnd:MicrosoftTheme
	}
}
