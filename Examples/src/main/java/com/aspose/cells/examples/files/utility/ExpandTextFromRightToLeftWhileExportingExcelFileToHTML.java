package com.aspose.cells.examples.files.utility;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ExpandTextFromRightToLeftWhileExportingExcelFileToHTML {

	public static void main(String[] args) throws Exception {
		
		// The path to the resource directory.
		String dataDir = Utils.getSharedDataDir(ExpandTextFromRightToLeftWhileExportingExcelFileToHTML.class) + "Conversion/";
		
		//Load source excel file inside the workbook object
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		//Save workbook in HTML format
		wb.save(dataDir + "output-" + CellsHelper.getVersion() + ".html", SaveFormat.HTML);
	}

}
