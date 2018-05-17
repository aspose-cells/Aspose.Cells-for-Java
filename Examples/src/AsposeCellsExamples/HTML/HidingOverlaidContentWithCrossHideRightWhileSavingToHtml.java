package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class HidingOverlaidContentWithCrossHideRightWhileSavingToHtml {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load sample Excel file 
		Workbook wb = new Workbook(srcDir + "sampleHidingOverlaidContentWithCrossHideRightWhileSavingToHtml.xlsx");

		//Specify HtmlSaveOptions - Hide Overlaid Content with CrossHideRight while saving to Html
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setHtmlCrossStringType(HtmlCrossType.CROSS_HIDE_RIGHT);

		//Save to HTML with HtmlSaveOptions
		wb.save(outDir + "outputHidingOverlaidContentWithCrossHideRightWhileSavingToHtml.html", opts);
	 
		// Print the message
		System.out.println("HidingOverlaidContentWithCrossHideRightWhileSavingToHtml executed successfully.");
	}
}
