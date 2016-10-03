package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class HowToCreatePieChart {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(HowToCreatePieChart.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the first worksheet
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// Adding some sample value to cells
		Cells cells = sheet.getCells();
		Cell cell = cells.get("A1");
		cell.setValue("Italy");
		cell = cells.get("A2");
		cell.setValue("Germany");
		cell = cells.get("A3");
		cell.setValue("England");
		cell = cells.get("A4");
		cell.setValue("Sweeden");
		cell = cells.get("A5");
		cell.setValue("America");
		cell = cells.get("A6");
		cell.setValue("London");
		cell = cells.get("A7");
		cell.setValue("Spain");
		cell = cells.get("A8");
		cell.setValue("France");

		cell = cells.get("B1");
		cell.setValue(10000);
		cell = cells.get("B2");
		cell.setValue(20000);
		cell = cells.get("B3");
		cell.setValue(45000);
		cell = cells.get("B4");
		cell.setValue(70000);
		cell = cells.get("B5");
		cell.setValue(19000);
		cell = cells.get("B6");
		cell.setValue(35000);
		cell = cells.get("B7");
		cell.setValue(28000);
		cell = cells.get("B8");
		cell.setValue(55000);

		ChartCollection charts = sheet.getCharts();

		// Adding a chart to the worksheet
		int chartIndex = charts.add(ChartType.PIE, 15, 4, 40, 15);
		Chart chart = charts.get(chartIndex);

		// Adding NSeries (chart data source) to the chart ranging from "A1"
		// cell to "B3"
		SeriesCollection serieses = chart.getNSeries();
		serieses.add("B1:B8", true);
		serieses.setCategoryData("A1:A8");
		serieses.setColorVaried(true);
		chart.setShowDataTable(true);

		// setting Title
		chart.getTitle().setText("Sales By Region");
		chart.getTitle().getFont().setColor(Color.getBlue());
		chart.getTitle().getFont().setBold(true);
		chart.getTitle().getFont().setSize(12);

		// Datalabels
		DataLabels datalabels;
		for (int i = 0; i < serieses.getCount(); i++) {
			datalabels = serieses.get(i).getDataLabels();
			datalabels.setPosition(LabelPositionType.INSIDE_BASE);
			datalabels.setShowCategoryName(true);
			datalabels.setShowValue(true);
			datalabels.setShowPercentage(false);
			datalabels.setShowLegendKey(true);
		}
		// Saving the Excel file
		workbook.save(dataDir + "HTCPChart_out.xls");

		// Print message
		System.out.println("Pie chart is successfully created.");


	}
}
