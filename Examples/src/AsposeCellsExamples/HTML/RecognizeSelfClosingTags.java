package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RecognizeSelfClosingTags {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // Set Html load options and keep precision true
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.HTML);

        // Load sample source file
        Workbook wb = new Workbook(srcDir + "sampleSelfClosingTags.html", loadOptions);

        // Save the workbook
        wb.save(outDir + "outsampleSelfClosingTags.xlsx");
        // ExEnd:1
		
		// Print the message
		System.out.println("RecognizeSelfClosingTags executed successfully.");
	}
}
