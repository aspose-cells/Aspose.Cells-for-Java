package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import AsposeCellsExamples.Utils;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

public class ConvertExcelFileToHtmlWithTooltip {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		//directories
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		// Open the template file
		Workbook workbook = new Workbook(sourceDir + "AddTooltipToHtmlSample.xlsx");

		HtmlSaveOptions options = new HtmlSaveOptions();
		options.setAddTooltipText(true);

		// Save as Markdown
		workbook.save(outputDir + "AddTooltipToHtmlSample_out.html", options);
        // ExEnd:1

		System.out.println("ConvertExcelFileToHtmlWithTooltip executed successfully.");
	}
}
