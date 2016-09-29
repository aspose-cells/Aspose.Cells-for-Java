package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.LegendPositionType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Series;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CustomSliceOrSectorColorsPieChart {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CustomSliceOrSectorColorsPieChart.class) + "articles/";
		// Create a workbook object from the template file
		Workbook workbook = new Workbook();

		// Access the first worksheet.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Put the sample values used in a pie chart
		worksheet.getCells().get("C3").putValue("India");
		worksheet.getCells().get("C4").putValue("China");
		worksheet.getCells().get("C5").putValue("United States");
		worksheet.getCells().get("C6").putValue("Russia");
		worksheet.getCells().get("C7").putValue("United Kingdom");
		worksheet.getCells().get("C8").putValue("Others");

		// Put the sample values used in a pie chart
		worksheet.getCells().get("D2").putValue("% of world population");
		worksheet.getCells().get("D3").putValue(25);
		worksheet.getCells().get("D4").putValue(30);
		worksheet.getCells().get("D5").putValue(10);
		worksheet.getCells().get("D6").putValue(13);
		worksheet.getCells().get("D7").putValue(9);
		worksheet.getCells().get("D8").putValue(13);

		// Create a pie chart with desired length and width
		int pieIdx = worksheet.getCharts().add(ChartType.PIE, 1, 6, 15, 14);

		// Access the pie chart
		Chart pie = worksheet.getCharts().get(pieIdx);

		// Set the pie chart series
		pie.getNSeries().add("D3:D8", true);

		// Set the category data
		pie.getNSeries().setCategoryData("=Sheet1!$C$3:$C$8");

		// Set the chart title that is linked to cell D2
		pie.getTitle().setLinkedSource("D2");

		// Set the legend position at the bottom.
		pie.getLegend().setPosition(LegendPositionType.BOTTOM);

		// Set the chart title's font name and color
		pie.getTitle().getFont().setName("Calibri");
		pie.getTitle().getFont().setSize(18);

		// Access the chart series
		Series srs = pie.getNSeries().get(0);

		// Color the indvidual points with custom colors
		srs.getPoints().get(0).getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0, 246, 22, 219));
		srs.getPoints().get(1).getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0, 51, 34, 84));
		srs.getPoints().get(2).getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0, 46, 74, 44));
		srs.getPoints().get(3).getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0, 19, 99, 44));
		srs.getPoints().get(4).getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0, 208, 223, 7));
		srs.getPoints().get(5).getArea().setForegroundColor(com.aspose.cells.Color.fromArgb(0, 222, 69, 8));

		// Autofit all columns
		worksheet.autoFitColumns();

		// Save the workbook
		workbook.save(dataDir + "CSOrSColorsPieChart_out.xlsx", SaveFormat.XLSX);

	}
}
