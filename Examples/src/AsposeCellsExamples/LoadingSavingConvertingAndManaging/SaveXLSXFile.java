package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class SaveXLSXFile {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveXLSXFile.class) + "LoadingSavingConvertingAndManaging/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in xlsx format
		workbook.save(dataDir + "SXLSXFile_out.xlsx", FileFormatType.XLSX);

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
