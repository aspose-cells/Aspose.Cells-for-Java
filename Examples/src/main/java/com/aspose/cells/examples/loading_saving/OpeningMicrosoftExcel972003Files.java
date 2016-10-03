package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class OpeningMicrosoftExcel972003Files {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningMicrosoftExcel972003Files.class) + "loading_saving/";

		// Opening Microsoft Excel 97 Files
		// Createing and EXCEL_97_TO_2003 LoadOptions object
		LoadOptions loadOptions1 = new LoadOptions(FileFormatType.EXCEL_97_TO_2003);

		// Creating an Workbook object with excel 97 file path and the
		// loadOptions object
		Workbook workbook3 = new Workbook(dataDir + "Book_Excel97_2003.xls", loadOptions1);

		// Print message
		System.out.println("Excel 97 Workbook opened successfully.");



	}
}
