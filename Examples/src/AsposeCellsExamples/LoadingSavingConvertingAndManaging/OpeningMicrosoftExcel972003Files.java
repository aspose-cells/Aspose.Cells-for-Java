package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class OpeningMicrosoftExcel972003Files {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningMicrosoftExcel972003Files.class) + "LoadingSavingConvertingAndManaging/";

		// Opening Microsoft Excel 97 Files
		// Creating an Workbook object with excel 97 file path
		new Workbook(dataDir + "Book_Excel97_2003.xls");

		// Print message
		System.out.println("Excel 97 Workbook opened successfully.");
	}
}
