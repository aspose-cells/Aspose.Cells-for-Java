package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import java.io.FileOutputStream;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class SavingFiletoStream {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingFiletoStream.class) + "LoadingSavingConvertingAndManaging/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		FileOutputStream stream = new FileOutputStream(dataDir + "SFToStream_out.xlsx");
		workbook.save(stream, SaveFormat.XLSX);

		// Print Message
		System.out.println("Worksheets are saved successfully.");
		stream.close();

	}
}
