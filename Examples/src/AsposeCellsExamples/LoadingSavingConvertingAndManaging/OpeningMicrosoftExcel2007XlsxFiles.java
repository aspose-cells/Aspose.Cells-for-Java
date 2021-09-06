package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class OpeningMicrosoftExcel2007XlsxFiles {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningMicrosoftExcel2007XlsxFiles.class) + "LoadingSavingConvertingAndManaging/";

		// Creating an Workbook object with 2007 xlsx file path
		Workbook workbook4 = new Workbook(dataDir + "Book_Excel2007.xlsx");

		System.out.println(workbook4.getFileName());
		// Print message
		System.out.println("Excel 2007 Workbook opened successfully.");
	}
}
