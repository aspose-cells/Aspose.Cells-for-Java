package com.aspose.cells.examples.SmartMarkers;

import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class UsingVariableArray {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingVariableArray.class) + "SmartMarkers/";
		// Instantiate a new Workbook designer.
		WorkbookDesigner report = new WorkbookDesigner();

		// Get the first worksheet of the workbook.
		Worksheet w = report.getWorkbook().getWorksheets().get(0);

		/*
		 * Set the Variable Array marker to a cell.You may also place this Smart
		 * Marker into a template file manually in Ms Excel and then open this
		 * file via Workbook.
		 */
		w.getCells().get("A1").putValue("&=$VariableArray");

		// Set the DataSource for the marker(s).
		report.setDataSource("VariableArray", new String[] { "English", "Arabic", "Hindi", "Urdu", "French" });

		// Process the markers.
		report.process(false);

		// Save the Excel file.
		report.getWorkbook().save(dataDir + "varaiblearray-out.xlsx");
	}
}
