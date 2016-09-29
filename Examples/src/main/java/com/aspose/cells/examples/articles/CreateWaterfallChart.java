package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.FormattingType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CreateWaterfallChart {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateWaterfallChart.class) + "articles/";

		// Create an instance of Workbook
		Workbook workbook = new Workbook();

		// Retrieve the first Worksheet in Workbook
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Retrieve the Cells of the first Worksheet
		Cells cells = worksheet.getCells();

		// Input some data which chart will use as source
		cells.get("A1").putValue("Previous Year");
		cells.get("A2").putValue("January");
		cells.get("A3").putValue("March");
		cells.get("A4").putValue("August");
		cells.get("A5").putValue("October");
		cells.get("A6").putValue("Current Year");

		cells.get("B1").putValue(8.5);
		cells.get("B2").putValue(1.5);
		cells.get("B3").putValue(7.5);
		cells.get("B4").putValue(7.5);
		cells.get("B5").putValue(8.5);
		cells.get("B6").putValue(3.5);

		cells.get("C1").putValue(1.5);
		cells.get("C2").putValue(4.5);
		cells.get("C3").putValue(3.5);
		cells.get("C4").putValue(9.5);
		cells.get("C5").putValue(7.5);
		cells.get("C6").putValue(9.5);

		// Add a Chart of type Line in same worksheet as of data
		int idx = worksheet.getCharts().add(ChartType.LINE, 4, 4, 25, 13);
		// Reterieve the Chart object
		Chart chart = worksheet.getCharts().get(idx);

		// Add Series
		chart.getNSeries().add("$B$1:$C$6", true);

		// Add Category Data
		chart.getNSeries().setCategoryData("$A$1:$A$6");

		// Series has Up Down Bars
		chart.getNSeries().get(0).setHasUpDownBars(true);

		// Set the colors of Up and Down Bars
		chart.getNSeries().get(0).getUpBars().getArea().setForegroundColor(Color.getGreen());
		chart.getNSeries().get(0).getDownBars().getArea().setForegroundColor(Color.getRed());

		// Make both Series Lines invisible
		chart.getNSeries().get(0).getBorder().setVisible(false);
		chart.getNSeries().get(1).getBorder().setVisible(false);

		// Set the Plot Area Formatting Automatic
		chart.getPlotArea().getArea().setFormatting(FormattingType.AUTOMATIC);

		// Delete the Legend
		chart.getLegend().getLegendEntries().get(0).setDeleted(true);
		chart.getLegend().getLegendEntries().get(1).setDeleted(true);

		// Save the workbook
		workbook.save(dataDir + "CWfallChart_out.xlsx");

	}
}
