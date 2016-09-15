package com.aspose.cells.examples.articles;

import com.aspose.cells.FileFormatInfo;
import com.aspose.cells.FileFormatUtil;
import com.aspose.cells.examples.Utils;

public class DetectFileFormatandCheckFileEncrypted {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DetectFileFormatandCheckFileEncrypted.class) + "articles/";

		// Detect file format
		FileFormatInfo info = FileFormatUtil.detectFileFormat(dataDir + "Book1.xlsx");

		// Gets the detected load format
		System.out.println("The spreadsheet format is: " + FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));

		// Check if the file is encrypted.
		System.out.println("The file is encrypted: " + info.isEncrypted());

	}
}
