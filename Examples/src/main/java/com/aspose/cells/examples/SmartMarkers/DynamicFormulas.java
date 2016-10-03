package com.aspose.cells.examples.SmartMarkers;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.examples.Utils;

public class DynamicFormulas {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DynamicFormulas.class) + "SmartMarkers/";
		// Instantiating a WorkbookDesigner object
		WorkbookDesigner designer = new WorkbookDesigner();

		// Set workbook which containing smart markers
		Workbook workbook = new Workbook(dataDir + "TestSmartMarkers.xlsx");
		designer.setWorkbook(workbook);

		// Set the data source for the designer spreadsheet
		designer.setDataSource(dataDir, workbook);

		// Process the smart markers
		designer.process();
	}
}