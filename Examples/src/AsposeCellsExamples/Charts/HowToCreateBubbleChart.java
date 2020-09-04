package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class HowToCreateBubbleChart {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(HowToCreateBubbleChart.class) + "Charts/";

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
		cell.setValue(180);
		cell = cells.get("C1");
		cell.setValue(320);
		cell = cells.get("C2");
		cell.setValue(110);
		cell = cells.get("C3");
		cell.setValue(180);
		cell = cells.get("D1");
		cell.setValue(40);
		cell = cells.get("D2");
		cell.setValue(120);
		cell = cells.get("D3");
		cell.setValue(250);

		ChartCollection charts = sheet.getCharts();

		// Adding a chart to the worksheet
		int chartIndex = charts.add(ChartType.BUBBLE, 5, 0, 15, 5);
		Chart chart = charts.get(chartIndex);

		// Adding NSeries (chart data source) to the chart ranging from "A1"
		// cell to "B3"
		SeriesCollection serieses = chart.getNSeries();
		serieses.add("A1:B3", true);

		// Set bubble sizes
		chart.getNSeries().get(0).setBubbleSizes("B2:D2");
		chart.getNSeries().get(0).setXValues("B3:D3");
		chart.getNSeries().get(0).setValues("B1:D1");

		// Saving the Excel file
		workbook.save(dataDir + "HToCrBChart_out.xls");
		// ExEnd:1
		
		// Print message
		System.out.println("Bubble chart is successfully created.");
	}
}
