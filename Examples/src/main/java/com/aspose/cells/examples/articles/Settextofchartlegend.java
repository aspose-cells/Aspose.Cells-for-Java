package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class Settextofchartlegend {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(Settextofchartlegend.class) + "articles/";

		// Open the template file.
		Workbook workbook = new Workbook(dataDir + "sample.xlsx");

		// Access the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Access the first chart inside the sheet
		Chart chart = sheet.getCharts().get(0);

		// Set text of second legend entry fill to none
		chart.getLegend().getLegendEntries().get(1).setTextNoFill(true);

		// Save the workbook in xlsx format
		workbook.save(dataDir + "Settextofchartlegend_out.xlsx", SaveFormat.XLSX);

	}

}
