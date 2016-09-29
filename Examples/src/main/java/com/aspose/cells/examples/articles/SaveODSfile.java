package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.OdsSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SaveODSfile {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveODSfile.class) + "articles/";
		// Create workbook
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Put some value in cell A1
		Cell cell = worksheet.getCells().get("A1");
		cell.putValue("Welcome to Aspose!");

		// Save ODS in ODF 1.2 version which is default
		OdsSaveOptions options = new OdsSaveOptions();
		workbook.save("SaveODSfile1_out.ods", options);

		// Save ODS in ODF 1.1 version
		options.setStrictSchema11(true);
		workbook.save("SaveODSfile2_out.ods", options);


	}
}
