package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class RenderCustomDateFormat {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the directories.
		String sourceDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "sampleRenderCustomDateFormat.xlsx");
		workbook.save(outDir + "sampleRenderCustomDateFormat_out.pdf");
		// ExEnd:1
	}
}
