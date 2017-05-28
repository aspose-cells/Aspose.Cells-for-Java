package AsposeCellsExamples.Charts;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ChangeChartPosition {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeChartPosition.class) + "Charts/";

		String filePath = dataDir + "chart.xls";

		Workbook workbook = new Workbook(filePath);

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Reposition the chart
		chart.getChartObject().setX(250);
		chart.getChartObject().setY(150);

		// Output the file
		workbook.save(dataDir + "CCPosition_out.xls");

		// Print message
		System.out.println("Position and Size of Chart is changed successfully.");

	}
}
