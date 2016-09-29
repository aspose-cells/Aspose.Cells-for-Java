package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.DataLabels;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ShowCellRangeAsTheDataLabels {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ShowCellRangeAsTheDataLabels.class) + "articles/";
		// Create workbook from the source Excel file
		Workbook workbook = new Workbook(dataDir + "sample.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the chart inside the worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Check the "Label Contains - Value From Cells"
		DataLabels dataLabels = chart.getNSeries().get(0).getDataLabels();
		dataLabels.setShowCellRange(true);

		// Save the workbook
		workbook.save(dataDir + "SCellRAsTheDataLabels_out.xlsx");


	}
}
