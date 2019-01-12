package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

import java.io.FileInputStream;

public class OpeningFilesThroughStream {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningFilesThroughStream.class) + "LoadingSavingConvertingAndManaging/";

		// Opening workbook from stream
		// Create a Stream object
		FileInputStream fstream = new FileInputStream(dataDir + "Book2.xls");

		// Creating an Workbook object with the stream object
		new Workbook(fstream);

		fstream.close();

		// Print message
		System.out.println("Workbook opened using stream successfully.");

	}
}
