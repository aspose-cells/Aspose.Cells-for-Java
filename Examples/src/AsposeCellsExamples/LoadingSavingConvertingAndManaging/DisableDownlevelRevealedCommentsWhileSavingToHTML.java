package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class DisableDownlevelRevealedCommentsWhileSavingToHTML {

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Load sample workbook
		Workbook wb = new Workbook(srcDir + "sampleDisableDownlevelRevealedComments.xlsx");
		 
		// Disable DisableDownlevelRevealedComments
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setDisableDownlevelRevealedComments(true);
		 
		// Save the workbook in html
		wb.save(outDir + "outputDisableDownlevelRevealedComments_true.html", opts);

		// Print the message
		System.out.println("DisableDownlevelRevealedCommentsWhileSavingToHTML executed successfully.");
	}
}
