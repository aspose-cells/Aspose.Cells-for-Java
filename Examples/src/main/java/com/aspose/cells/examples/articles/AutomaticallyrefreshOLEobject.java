package com.aspose.cells.examples.articles;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AutomaticallyrefreshOLEobject {
	public static void main(String[] args) throws Exception {
		// ExStart:AutomaticallyrefreshOLEobject
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AutomaticallyrefreshOLEobject.class);

		// Create workbook object from your sample excel file
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		// Access first worksheet
		Worksheet sheet = wb.getWorksheets().get(0);

		// Set auto load property of first ole object to true
		sheet.getOleObjects().get(0).setAutoLoad(true);

		// Save the worbook in xlsx format
		wb.save(dataDir + "output.xlsx", SaveFormat.XLSX);
		// ExEnd:AutomaticallyrefreshOLEobject
	}

}
