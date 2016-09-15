package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.Trendline;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GetEquationText {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetEquationText.class) + "articles/";
		// Create workbook object from source Excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first chart inside the worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Calculate the Chart first to get the Equation Text of Trendline
		chart.calculate();

		// Access the Trendline
		Trendline trendLine = chart.getNSeries().get(0).getTrendLines().get(0);

		// Read the Equation Text of Trendline
		System.out.println("Equation Text: " + trendLine.getDataLabels().getText());

	}
}
