package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class HandleAutomaticUnitsOfChartAxisLikeMicrosoftExcel {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		// ExStart:1
		//Load the sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleHandleAutomaticUnitsOfChartAxisLikeMicrosoftExcel.xlsx");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access first chart
		Chart ch = ws.getCharts().get(0);

		//Render chart to pdf
		ch.toPdf(outDir + "outputHandleAutomaticUnitsOfChartAxisLikeMicrosoftExcel.pdf");
		
		// Print the message
		System.out.println("HandleAutomaticUnitsOfChartAxisLikeMicrosoftExcel executed successfully.");
		// ExEnd:1
	}
}
