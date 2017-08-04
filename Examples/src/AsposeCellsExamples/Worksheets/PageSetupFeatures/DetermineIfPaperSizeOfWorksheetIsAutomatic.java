package AsposeCellsExamples.Worksheets.PageSetupFeatures;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class DetermineIfPaperSizeOfWorksheetIsAutomatic { 

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Load the first workbook having automatic paper size false
		Workbook wb1 = new Workbook(srcDir + "samplePageSetupIsAutomaticPaperSize-False.xlsx");

		// Load the second workbook having automatic paper size true
		Workbook wb2 = new Workbook(srcDir + "samplePageSetupIsAutomaticPaperSize-True.xlsx");

		// Access first worksheet of both workbooks
		Worksheet ws11 = wb1.getWorksheets().get(0);
		Worksheet ws12 = wb2.getWorksheets().get(0);

		// Print the PageSetup.IsAutomaticPaperSize property of both worksheets
		System.out.println("First Worksheet of First Workbook - IsAutomaticPaperSize: "
				+ ws11.getPageSetup().isAutomaticPaperSize());
		System.out.println("First Worksheet of Second Workbook - IsAutomaticPaperSize: "
				+ ws12.getPageSetup().isAutomaticPaperSize());

		// Print the message
		System.out.println("DetermineIfPaperSizeOfWorksheetIsAutomatic executed successfully.");
	}
}
