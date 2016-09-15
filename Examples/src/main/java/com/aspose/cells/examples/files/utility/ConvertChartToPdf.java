package com.aspose.cells.examples.files.utility;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertChartToPdf {

	public static void main(String[] args) throws Exception {


		String dataDir = Utils.getSharedDataDir(ConvertChartToPdf.class) + "files/";
		String inputPath = dataDir + "Sample1.xls";
		String outputPath = dataDir + "CChartToPdf.pdf";

		// Load excel file containing charts
		Workbook workbook = new Workbook(inputPath);

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first chart inside the worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Save the chart into pdf format
		chart.toPdf(outputPath);

		System.out.println("File saved " + outputPath);

	}
}
