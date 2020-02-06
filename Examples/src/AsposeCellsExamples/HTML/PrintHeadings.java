package AsposeCellsExamples.HTML;

import AsposeCellsExamples.Utils;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;

public class PrintHeadings {
	
	static String sourceDir = Utils.Get_SourceDirectory();
	static String outputDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
        // ExStart:1
		// Load the Excel file.
        Workbook workbook = new Workbook(sourceDir + "Book1.xlsx");

        // Initialize HtmlSaveOptions
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportHeadings(true);

        //Save to HTML format
        workbook.save(outputDir + "PrintHeadings_out.html", options);
        // ExEnd:1
        
		// Print the message
		System.out.println("PrintHeadings executed successfully.");
	}
}
