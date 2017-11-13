package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyHtmlCrossTypeInOutputHTML {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Enum to String
		String[] strsHtmlCrossStringType = new String[]{"Default", "MSExport", "Cross", "FitToCell"};
		 
		//Load the sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleHtmlCrossStringType.xlsx");
		 
		//Specify HTML Cross Type
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setHtmlCrossStringType(HtmlCrossType.DEFAULT);
		opts.setHtmlCrossStringType(HtmlCrossType.MS_EXPORT);
		opts.setHtmlCrossStringType(HtmlCrossType.CROSS);
		opts.setHtmlCrossStringType(HtmlCrossType.FIT_TO_CELL);
		 
		//Output Html
		String strHtmlCrossStringType = strsHtmlCrossStringType[opts.getHtmlCrossStringType()];
		wb.save(outDir + "out" + strHtmlCrossStringType + ".htm", opts);

		// Print the message
		System.out.println("SpecifyHtmlCrossTypeInOutputHTML executed successfully.");
	}
}
