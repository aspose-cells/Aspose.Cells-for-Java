package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CreateChart {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateChart.class) + "Charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the first worksheet
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// Adding some sample value to cells
		Cells cells = sheet.getCells();
		Cell cell = cells.get("A1");
		cell.setValue(50);
		cell = cells.get("A2");
		cell.setValue(100);
		cell = cells.get("A3");
		cell.setValue(150);
		cell = cells.get("B1");
		cell.setValue(4);
		cell = cells.get("B2");
		cell.setValue(20);
		cell = cells.get("B3");
		cell.setValue(50);

		ChartCollection charts = sheet.getCharts();

		// Adding a chart to the worksheet
		int chartIndex = charts.add(ChartType.PYRAMID, 5, 0, 15, 5);
		Chart chart = charts.get(chartIndex);

		// Adding NSeries (chart data source) to the chart ranging from "A1"
		// cell to "B3"
		SeriesCollection serieses = chart.getNSeries();
		serieses.add("A1:B3", true);

		// Saving the Excel file
		workbook.save(dataDir + "CreateChart_out.xls");
		// ExEnd:1
		
		// Print message
		System.out.println("Workbook with chart is successfully created.");
	}
}
