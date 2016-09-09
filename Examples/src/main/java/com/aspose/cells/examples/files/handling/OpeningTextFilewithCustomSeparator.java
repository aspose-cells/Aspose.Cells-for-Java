package com.aspose.cells.examples.files.handling;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class OpeningTextFilewithCustomSeparator {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningTextFilewithCustomSeparator.class) + "files/";
		String filePath = dataDir + "Book11.csv";

		TxtLoadOptions txtoption = new TxtLoadOptions();
		txtoption.setSeparator(',');
		txtoption.setEncoding(Encoding.getUTF8());

		// Creating Workbook object and saving it
		Workbook workbook = new Workbook(dataDir + "Book11.csv", txtoption);
		workbook.save(dataDir + "OTFWCSeparator-out.pdf", FileFormatType.PDF);

		// Print message
		System.out.println("Custom Separator workbook has been opened successfully.");
		// ExEnd:1

	}
}
