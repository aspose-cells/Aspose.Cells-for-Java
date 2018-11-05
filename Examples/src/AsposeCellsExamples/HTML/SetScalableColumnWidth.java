package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SetScalableColumnWidth {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // Load sample source file
        Workbook wb = new Workbook(srcDir + "sampleForScalableColumns.xlsx");

        // Specify Html Save Options
        HtmlSaveOptions options = new HtmlSaveOptions();

        // Set the property for scalable width
        options.setWidthScalable(true);

        // Specify image save format
        options.setExportImagesAsBase64(true);

        // Save the workbook in Html format with specified Html Save Options
        wb.save(outDir + "outsampleForScalableColumns.html", options);
        // ExEnd:1
		
		// Print the message
		System.out.println("SetScalableColumnWidth executed successfully.");
	}
}
