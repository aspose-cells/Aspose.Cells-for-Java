package AsposeCellsExamples.Charts;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class ChangeTickLabelDirection {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		//Source directory
		String sourceDir = Utils.Get_SourceDirectory();

		//Output directory
		String outputDir = Utils.Get_OutputDirectory();

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(sourceDir + "SampleChangeTickLabelDirection.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		chart.getCategoryAxis().getTickLabels().setDirectionType(ChartTextDirectionType.HORIZONTAL);

		// Output the file
		workbook.save(outputDir + "outputChangeTickLabelDirection.xlsx");
		// ExEnd:1
		
		// Print message
		System.out.println("ChangeTickLabelDirection executed successfully.");
	}
}
