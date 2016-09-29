package com.aspose.cells.examples.articles;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ConvertXLSFileToPDF {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertXLSFileToPDF.class) + "articles/";
		
		//Create a new Workbook
		Workbook book = new Workbook(dataDir + "SampleInput.xlsx");

		//Save the excel file to PDF format
		book.save(dataDir + "CXLSFileToPDF_out.pdf", SaveFormat.PDF);

	}
}
