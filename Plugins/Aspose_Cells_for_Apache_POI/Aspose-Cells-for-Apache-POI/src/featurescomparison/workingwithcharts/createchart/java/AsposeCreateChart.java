package featurescomparison.workingwithcharts.createchart.java;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.Series;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeCreateChart
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcharts/createchart/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Obtaining the reference of the newly added worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		//Adding a sample value to "A1" cell
		cells.get("A1").setValue(50);

		//Adding a sample value to "A2" cell
		cells.get("A2").setValue(100);

		//Adding a sample value to "A3" cell
		cells.get("A3").setValue(150);

		//Adding a sample value to "A4" cell
		cells.get("A4").setValue(200);

		//Adding a sample value to "B1" cell
		cells.get("B1").setValue(60);

		//Adding a sample value to "B2" cell
		cells.get("B2").setValue(32);

		//Adding a sample value to "B3" cell
		cells.get("B3").setValue(50);

		//Adding a sample value to "B4" cell
		cells.get("B4").setValue(40);

		//Adding a chart to the worksheet
		ChartCollection charts = worksheet.getCharts();

		//Accessing the instance of the newly added chart
		int chartIndex = worksheet.getCharts().add(ChartType.COLUMN,5,0,15,5);
		Chart chart = worksheet.getCharts().get(chartIndex);

		//Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B4"
		SeriesCollection nSeries = chart.getNSeries();
		nSeries.add("A1:B4",true);

		//Setting the chart type of 2nd NSeries to display as line chart
		Series series = nSeries.get(1);
		series.setType(ChartType.LINE);

		//Saving the Excel file
		workbook.save(dataPath + "AsposeChart.xls");
		
		System.out.println("Done.");
	}
}
