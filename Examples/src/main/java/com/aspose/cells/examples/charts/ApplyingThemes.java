package com.aspose.cells.examples.charts;

import com.aspose.cells.CellsColor;
import com.aspose.cells.Chart;
import com.aspose.cells.FillType;
import com.aspose.cells.ThemeColor;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ApplyingThemes {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ApplyingThemes.class) + "charts/";

		// Instantiate the workbook to open the file that contains a chart
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Get the first chart in the sheet
		Chart chart = worksheet.getCharts().get(0);

		// Specify the FilFormat's type to Solid Fill of the first series
		chart.getNSeries().get(0).getArea().getFillFormat().setType(FillType.SOLID);

		// Get the CellsColor of SolidFill
		CellsColor cc = chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().getCellsColor();

		// Create a theme in Accent style
		cc.setThemeColor(new ThemeColor(ThemeColorType.ACCENT_6, 0.6));

		// Apply the them to the series
		chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().setCellsColor(cc);

		// Save the Excel file
		workbook.save(dataDir + "AThemes_out.xlsx");
	}
}
