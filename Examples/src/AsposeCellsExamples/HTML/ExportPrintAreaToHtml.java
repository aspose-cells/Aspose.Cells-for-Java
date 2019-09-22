package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExportPrintAreaToHtml {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
        // ExStart:1
		// Load the Excel file.
        Workbook wb = new Workbook(srcDir + "sampleInlineCharts.xlsx");

        // Access the sheet
        Worksheet ws = wb.getWorksheets().get(0);

        // Set the print area.
        ws.getPageSetup().setPrintArea("D2:M20");

        // Initialize HtmlSaveOptions
        HtmlSaveOptions options = new HtmlSaveOptions();

        // Set flag to export print area only
        options.setExportPrintAreaOnly(true);

        //Save to HTML format
        wb.save(outDir + "outputInlineCharts.html",options);
        // ExEnd:1
        
		// Print the message
		System.out.println("ExportPrintAreaToHtml executed successfully.");
	}
}
