package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DisableTextWrapping {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DisableTextWrapping.class) + "articles/";
		// Load the sample Excel file inside the workbook object
		Workbook workbook = new Workbook(dataDir + "SampleChart.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first chart inside the worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Disable the Text Wrapping of Data Labels in all Series
		chart.getNSeries().get(0).getDataLabels().setTextWrapped(false);
		chart.getNSeries().get(1).getDataLabels().setTextWrapped(false);
		chart.getNSeries().get(2).getDataLabels().setTextWrapped(false);

		// Save the workbook
		workbook.save(dataDir + "DTextWrapping_out.xlsx");

	}
}
