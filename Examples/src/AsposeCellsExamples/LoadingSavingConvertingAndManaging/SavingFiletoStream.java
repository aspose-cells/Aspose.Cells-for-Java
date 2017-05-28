package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.XpsSaveOptions;
import AsposeCellsExamples.Utils;

import java.io.FileOutputStream;

public class SavingFiletoStream {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingFiletoStream.class) + "LoadingSavingConvertingAndManaging/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		FileOutputStream stream = new FileOutputStream(dataDir + "SFToStream_out.xlsx");
		workbook.save(stream, new XpsSaveOptions(FileFormatType.XLSX));

		// Print Message
		System.out.println("Worksheets are saved successfully.");
		stream.close();

	}
}
