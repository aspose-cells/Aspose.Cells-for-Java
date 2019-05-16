package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetChartSubTitleForODSFile {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetChartSubTitleForODSFile.class) + "Charts/";

		String filePath = dataDir + "SampleChart.ods";

		Workbook workbook = new Workbook(filePath);

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		System.out.println("Chart Subtitle: " + chart.getSubTitle().getText());
        // ExEnd:1

		System.out.println("GetChartSubTitleForODSFile executed successfully.");
	}
}
