package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.Range;
import com.aspose.cells.SheetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GenerateChartByProcessingSmartMarkers {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GenerateChartByProcessingSmartMarkers.class) + "articles/";



		// Create an instance of Workbook
		Workbook book = new Workbook();

		// Access the first (default) Worksheet from the collection
		Worksheet dataSheet = book.getWorksheets().get(0);

		// Name the first Worksheet for referencing
		dataSheet.setName("ChartData");

		// Access the CellsCollection of ChartData Worksheet
		Cells cells = dataSheet.getCells();

		// Place the markers in the Worksheet according to desired layout
		cells.get("A1").putValue("&=$Headers(horizontal)");
		cells.get("A2").putValue("&=$Year2000(horizontal)");
		cells.get("A3").putValue("&=$Year2005(horizontal)");
		cells.get("A4").putValue("&=$Year2010(horizontal)");
		cells.get("A5").putValue("&=$Year2015(horizontal)");





		// Create string arrays which will serve as data sources to the smart markers
		String[] headers = new String[] { "", "Item 1", "Item 2", "Item 3", "Item 4", "Item 5", "Item 6", "Item 7",
				"Item 8", "Item 9", "Item 10", "Item 11", "Item 12" };
		String[] year2000 = new String[] { "2000", "310", "0", "110", "15", "20", "25", "30", "1222", "200", "421",
				"210", "133" };
		String[] year2005 = new String[] { "2005", "508", "0", "170", "280", "190", "400", "105", "132", "303", "199",
				"120", "100" };
		String[] year2010 = new String[] { "2010", "0", "210", "230", "1420", "1530", "160", "170", "110", "199", "129",
				"120", "230" };
		String[] year2015 = new String[] { "2015", "2818", "320", "340", "260", "210", "310", "220", "0", "0", "0", "0",
				"122" };





		// Create an instance of WorkbookDesigner
		WorkbookDesigner designer = new WorkbookDesigner();

		// Set the Workbook property for the instance of WorkbookDesigner
		designer.setWorkbook(book);

		// Set data sources for smart markers
		designer.setDataSource("Headers", headers);
		designer.setDataSource("Year2000", year2000);
		designer.setDataSource("Year2005", year2005);
		designer.setDataSource("Year2010", year2010);
		designer.setDataSource("Year2015", year2015);

		// Process the designer spreadsheet against the provided data sources
		designer.process();



		// Convert all string values of ChartData to numbers
		// This is an additional step as we have imported the string values
		dataSheet.getCells().convertStringToNumericValue();

		// Save the number of rows & columns from the ChartData in separate variables
		// These values will be used later to identify the chart's data range from ChartData
		int chartRows = dataSheet.getCells().getMaxDataRow() + 1;
		int chartCols = dataSheet.getCells().getMaxDataColumn() + 1;

		// Add a new Worksheet of type Chart to Workbook
		int chartSheetIdx = book.getWorksheets().add(SheetType.CHART);

		// Access the newly added Worksheet via its index
		Worksheet chartSheet = book.getWorksheets().get(chartSheetIdx);

		// Name the Worksheet
		chartSheet.setName("Chart");

		// Add a chart of type ColumnStacked to newly added Worksheet
		int chartIdx = chartSheet.getCharts().add(ChartType.COLUMN_STACKED, 0, 0, chartRows, chartCols);

		// Access the newly added Chart via its index
		Chart chart = chartSheet.getCharts().get(chartIdx);

		// Identify the chart's data range
		Range dataRange = dataSheet.getCells().createRange(0, 1, chartRows, chartCols - 1);

		// Set the data range for the chart
		chart.setChartDataRange(dataRange.getRefersTo(), false);

		// Set the chart to size with window
		chart.setSizeWithWindow(true);

		// Set the format for the tick labels
		chart.getValueAxis().getTickLabels().setNumberFormat("$###,### K");

		// Set chart title
		chart.getTitle().setText("Sales Summary");

		// Set ChartSheet an active sheet
		book.getWorksheets().setActiveSheetIndex(chartSheetIdx);

		// Save the final result
		book.save(dataDir + "GCByPSmartMarkers.xlsx");

	}
}
