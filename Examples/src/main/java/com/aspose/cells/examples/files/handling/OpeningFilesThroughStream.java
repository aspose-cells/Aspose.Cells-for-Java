package com.aspose.cells.examples.files.handling;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

import java.io.FileInputStream;

public class OpeningFilesThroughStream {

	public static void main(String[] args) throws Exception {
		// ExStart:OpeningFilesThroughStream
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(OpeningFilesThroughStream.class);

		// Opening workbook from stream
		// Create a Stream object
		FileInputStream fstream = new FileInputStream(dataDir + "Book2.xls");

		// Creating an Workbook object with the stream object
		Workbook workbook2 = new Workbook(fstream);

		fstream.close();

		// Print message
		System.out.println("Workbook opened using stream successfully.");
		// ExEnd:OpeningFilesThroughStream
	}
}
