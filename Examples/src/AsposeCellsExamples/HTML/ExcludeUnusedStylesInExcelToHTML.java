package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExcludeUnusedStylesInExcelToHTML { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create workbook
		Workbook wb = new Workbook();

		//Create an unused named style
		wb.createStyle().setName("UnusedStyle_XXXXXXXXXXXXXX");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Put some value in cell C7
		ws.getCells().get("C7").putValue("This is sample text.");

		//Specify html save options, we want to exclude unused styles
		HtmlSaveOptions opts = new HtmlSaveOptions();

		//Comment this line to include unused styles
		opts.setExcludeUnusedStyles(true);

		//Save the workbook in html format
		wb.save(outDir + "outputExcludeUnusedStylesInExcelToHTML.html", opts);
		
		// Print the message
		System.out.println("ExcludeUnusedStylesInExcelToHTML executed successfully.");
	}
}
