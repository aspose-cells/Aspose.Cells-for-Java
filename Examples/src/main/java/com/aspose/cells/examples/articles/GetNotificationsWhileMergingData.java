package com.aspose.cells.examples.articles;

import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GetNotificationsWhileMergingData {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetNotificationsWhileMergingData.class) + "articles/";
		// Instantiate a new Workbook designer
		WorkbookDesigner report = new WorkbookDesigner();

		// Get the first worksheet of the workbook
		Worksheet sheet = report.getWorkbook().getWorksheets().get(0);

		/*
		 * Set the Variable Array marker to a cell. You may also place this Smart Marker into a template file manually using Excel
		 * and then open this file via WorkbookDesigner
		 */
		sheet.getCells().get("A1").putValue("&=$VariableArray");

		// Set the data source for the marker(s)
		report.setDataSource("VariableArray", new String[] { "English", "Arabic", "Hindi", "Urdu", "French" });

		// Set the CallBack property
		report.setCallBack(new SmartMarkerCallBack(report.getWorkbook()));

		// Process the markers
		report.process(false);

		// Save the result
		report.getWorkbook().save(dataDir);

	}
}
