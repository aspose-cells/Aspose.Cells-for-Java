package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ExportCharttoSVG {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportCharttoSVG.class) + "articles/";
		// Create workbook object from source file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first chart inside the worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Set image or print options
		// with SVGFitToViewPort true
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setSaveFormat(SaveFormat.SVG);
		opts.setSVGFitToViewPort(true);

		// Save the chart to svg format
		chart.toImage(dataDir + "ECharttoSVG_out.svg", opts);


	}
}
