package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CreateChartPDFWithDesiredPageSize {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load sample Excel file containing the chart.
		Workbook wb = new Workbook(srcDir + "sampleCreateChartPDFWithDesiredPageSize.xlsx");
		 
		//Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Access first chart inside the worksheet.
		Chart ch = ws.getCharts().get(0);
		 
		//Create chart pdf with desired page size.
		ch.toPdf(outDir + "outputCreateChartPDFWithDesiredPageSize.pdf", 7, 7, PageLayoutAlignmentType.CENTER, PageLayoutAlignmentType.CENTER);
			 
		// Print the message
		System.out.println("CreateChartPDFWithDesiredPageSize executed successfully.");
	}
}
