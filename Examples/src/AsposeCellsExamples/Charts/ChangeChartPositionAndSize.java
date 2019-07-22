package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ChangeChartPositionAndSize {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeChartPositionAndSize.class) + "Charts/";

		String filePath = dataDir + "book1.xls";

		Workbook workbook = new Workbook(filePath);

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Resize the chart
		chart.getChartObject().setWidth(400);
		chart.getChartObject().setHeight(300);

		// Reposition the chart
		chart.getChartObject().setX(250);
		chart.getChartObject().setY(150);

		// Output the file
		workbook.save(dataDir + "ChangeChartPositionAndSize_out.xls");
		// ExEnd:1
		
		// Print message
		System.out.println("Position and Size of Chart is changed successfully.");
	}
}
