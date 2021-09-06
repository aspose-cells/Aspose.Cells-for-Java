package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class SaveXLSFile {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveXLSFile.class) + "LoadingSavingConvertingAndManaging/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in xls format
		workbook.save(dataDir + "SXLSFile_out.xls");

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
