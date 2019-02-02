package AsposeCellsExamples.Rendering;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RenderUsingTextCrossType {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// Load the template file
        Workbook wb = new Workbook(srcDir + "sampleCrosssType.xlsx");

        // Initialize PDF save options
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        
        // Set text criss type
        saveOptions.setTextCrossType(TextCrossType.STRICT_IN_CELL);
        
        // Save output PDF file
        wb.save(outDir + "outputCrosssType.pdf", saveOptions);

		// Print the message
		System.out.println("RenderUsingTextCrossType executed successfully.");
	}
}
