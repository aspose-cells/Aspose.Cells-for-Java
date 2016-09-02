package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertCharttoImageinSVGFormat {
	public static void main(String[] args) throws Exception {
		// ExStart:ConvertCharttoImageinSVGFormat
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ConvertCharttoImageinSVGFormat.class);
		// Create workbook object from source Excel file
		Workbook workbook = new Workbook(dataDir + "sample.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first chart inside the worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Save the chart into image in SVG format
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		options.setSaveFormat(SaveFormat.SVG);
		chart.toImage(dataDir + "ChartImage.svg", options);
		// ExEnd:ConvertCharttoImageinSVGFormat
	}
}
