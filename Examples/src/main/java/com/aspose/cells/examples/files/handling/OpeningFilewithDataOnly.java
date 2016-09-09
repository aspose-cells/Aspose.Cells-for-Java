package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadDataOption;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class OpeningFilewithDataOnly {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningFilewithDataOnly.class) + "files/";
		// Opening CSV Files
		// Creating and CSV LoadOptions object
		LoadOptions loadOptions = new LoadOptions(FileFormatType.XLSX);

		LoadDataOption dataoption = new LoadDataOption();
		dataoption.SheetNames = new String[] { "Sheet2" };
		dataoption.setImportFormula(true);

		loadOptions.setLoadDataAndFormatting(false);

		// Create a Workbook object and opening the file from its path
		Workbook wb = new Workbook(dataDir + "Book1.xlsx", loadOptions);
		System.out.println("File data imported successfully!");
		// ExEnd:1

	}
}
