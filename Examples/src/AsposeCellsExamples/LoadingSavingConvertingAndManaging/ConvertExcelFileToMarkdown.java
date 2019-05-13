package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ConvertExcelFileToMarkdown {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertExcelFileToMarkdown.class) + "LoadingSavingConvertingAndManaging/";

		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Save as Markdown
        workbook.save(dataDir + "Book1.md", SaveFormat.MARKDOWN);
        // ExEnd:1

		System.out.println("ConvertExcelFileToMarkdown executed successfully.");
	}
}
