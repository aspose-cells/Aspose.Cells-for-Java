package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CustomTextforOtherLabelofPieChart {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CustomTextforOtherLabelofPieChart.class) + "articles/";
		
		//Loads an existing spreadsheet containing a pie chart
		Workbook book = new Workbook(dataDir + "sample.xlsx");

		//Assigns the GlobalizationSettings property of the WorkbookSettings class
		//to the class created in first step
		book.getSettings().setGlobalizationSettings(new CustomSettings());

		//Accesses the 1st worksheet from the collection which contains pie chart
		Worksheet sheet = book.getWorksheets().get(0);

		//Accesses the 1st chart from the collection
		Chart chart = sheet.getCharts().get(0);

		//Refreshes the chart
		chart.calculate();

		//Renders the chart to image
		chart.toImage(dataDir + "CustomTextforOtherLabelofPieChart_out.png", new ImageOrPrintOptions());
	}
}
