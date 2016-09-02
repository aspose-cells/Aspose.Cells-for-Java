package com.aspose.cells.examples.AdvancedTopics.SmartMarkers;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.examples.Utils;

public class DynamicFormulas {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(DynamicFormulas.class);
		// Instantiating a WorkbookDesigner object
		WorkbookDesigner designer = new WorkbookDesigner();

		// Set workbook which containing smart markers
		Workbook workbook = new Workbook(dataDir + "designerFile");
		designer.setWorkbook(workbook);

		// Set the data source for the designer spreadsheet
		designer.setDataSource(dataSet);

		// Process the smart markers
		designer.process();
	}
}
