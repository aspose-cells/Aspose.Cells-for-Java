package AsposeCellsExamples.Charts;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

import java.awt.*;

public class InsertCheckboxInChartSheet {
	public static void main(String[] args) throws Exception {
		// ExStart: 1
		// directories
		String outputDir = Utils.Get_OutputDirectory();

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a chart to the worksheet
		int index = workbook.getWorksheets().add(SheetType.CHART);

		Worksheet sheet = workbook.getWorksheets().get(index);
		sheet.getCharts().addFloatingChart(ChartType.COLUMN, 0, 0, 1024, 960);
		sheet.getCharts().get(0).getNSeries().add("{1,2,3}", false);

		// Add checkbox to the chart.
		sheet.getCharts().get(0).getShapes().addShapeInChart(MsoDrawingType.CHECK_BOX, PlacementType.MOVE, 400, 400, 1000, 600);
		sheet.getCharts().get(0).getShapes().get(0).setText("CheckBox 1");

		// Convert chart to image with additional settings
		workbook.save(outputDir + "InsertCheckboxInChartSheet_out.xlsx");
		// ExEnd:1

		System.out.println("InsertCheckboxInChartSheet executed successfully.");
	}
}
