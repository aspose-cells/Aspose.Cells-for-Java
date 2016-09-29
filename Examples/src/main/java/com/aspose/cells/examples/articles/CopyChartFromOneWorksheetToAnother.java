package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartShape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class CopyChartFromOneWorksheetToAnother {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyChartFromOneWorksheetToAnother.class) + "articles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Shapes.xls");

		WorksheetCollection ws = workbook.getWorksheets();
		Worksheet sheet1 = ws.get("Chart");
		Worksheet sheet2 = ws.get("Result");

		// get the Chart from first worksheet
		Chart chart = sheet1.getCharts().get(0);

		ChartShape cshape = chart.getChartObject();

		// Copy the Chart to Second Worksheet
		sheet2.getShapes().addCopy(cshape, 20, 0, 2, 0);

		// Save the workbook
		workbook.save(dataDir + "CCFOneWToAnother_out.xls");

	}
}
