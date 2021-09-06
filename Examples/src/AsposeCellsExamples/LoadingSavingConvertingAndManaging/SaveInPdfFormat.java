package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class SaveInPdfFormat {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveInPdfFormat.class) + "LoadingSavingConvertingAndManaging/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in PDF format
		workbook.save(dataDir + "SIPdfFormat_out.pdf");

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
