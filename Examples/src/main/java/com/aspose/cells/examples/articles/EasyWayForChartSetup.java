package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class EasyWayForChartSetup {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(EasyWayForChartSetup.class) + "articles/";
		// Create new workbook
		Workbook workbook = new Workbook(FileFormatType.XLSX);

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add data for chart

		// Category Axis Values
		worksheet.getCells().get("A2").putValue("C1");
		worksheet.getCells().get("A3").putValue("C2");
		worksheet.getCells().get("A4").putValue("C3");

		// First vertical series
		worksheet.getCells().get("B1").putValue("T1");
		worksheet.getCells().get("B2").putValue(6);
		worksheet.getCells().get("B3").putValue(3);
		worksheet.getCells().get("B4").putValue(2);

		// Second vertical series
		worksheet.getCells().get("C1").putValue("T2");
		worksheet.getCells().get("C2").putValue(7);
		worksheet.getCells().get("C3").putValue(2);
		worksheet.getCells().get("C4").putValue(5);

		// Third vertical series
		worksheet.getCells().get("D1").putValue("T3");
		worksheet.getCells().get("D2").putValue(8);
		worksheet.getCells().get("D3").putValue(4);
		worksheet.getCells().get("D4").putValue(2);

		// Create Column chart with easy way
		int idx = worksheet.getCharts().add(ChartType.COLUMN, 6, 5, 20, 13);
		Chart ch = worksheet.getCharts().get(idx);
		ch.setChartDataRange("A1:D4", true);

		// Save the workbook
		workbook.save(dataDir + "EWForChartSetup.xlsx", SaveFormat.XLSX);

	}
}
