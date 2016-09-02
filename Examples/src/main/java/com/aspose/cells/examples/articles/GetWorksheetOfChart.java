package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GetWorksheetOfChart {
	public static void main(String[] args) throws Exception {
		// ExStart:GetWorksheetOfChart
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(GetWorksheetOfChart.class);

		// Create workbook from sample Excel file
		Workbook workbook = new Workbook(dataDir + "sample.xlsx");

		// Access first worksheet of the workbook
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Print worksheet name
		System.out.println("Sheet Name: " + worksheet.getName());

		// Access the first chart inside this worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Access the chart's sheet and display its name again
		System.out.println("Chart's Sheet Name: " + chart.getWorksheet().getName());
		// ExEnd:GetWorksheetOfChart
	}
}
