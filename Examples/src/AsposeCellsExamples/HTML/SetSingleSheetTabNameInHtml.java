package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SetSingleSheetTabNameInHtml {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main() throws Exception {
        // ExStart:1
        // Load the sample Excel file containing single sheet only
        Workbook wb = new Workbook(srcDir + "sampleSingleSheet.xlsx");

        // Specify HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions();

        // Set optional settings if required
        options.setEncoding(Encoding.getUTF8());
        options.setExportImagesAsBase64(true);
        options.setExportGridLines(true);
        options.setExportSimilarBorderStyle(true);
        options.setExportBogusRowData(true);
        options.setExcludeUnusedStyles(true);
        options.setExportHiddenWorksheet(true);

        //Save the workbook in Html format with specified Html Save Options
        wb.save(outDir + "outputSampleSingleSheet.htm", options);
        // ExEnd:1

		// Print the message
		System.out.println("SetSingleSheetTabNameInHtml executed successfully.");
	}
}
