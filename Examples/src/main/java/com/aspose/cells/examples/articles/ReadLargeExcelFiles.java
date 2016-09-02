package com.aspose.cells.examples.articles;

import com.aspose.cells.LoadOptions;
import com.aspose.cells.MemorySetting;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ReadLargeExcelFiles {
	public static void main(String[] args) throws Exception {
		// ExStart:ReadLargeExcelFiles
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ReadLargeExcelFiles.class);
		// Specify the LoadOptions
		LoadOptions opt = new LoadOptions();
		// Set the memory preferences
		opt.setMemorySetting(MemorySetting.MEMORY_PREFERENCE);
		// Instantiate the Workbook
		// Load the Big Excel file having large Data set in it
		Workbook wb = new Workbook(dataDir + "Book1.xlsx", opt);
		// ExEnd:ReadLargeExcelFiles
	}
}
