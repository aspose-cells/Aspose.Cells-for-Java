package com.aspose.cells.examples.SmartMarkers;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.examples.Utils;

public class UsingHTMLProperty {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingHTMLProperty.class) + "SmartMarkers/";
		Workbook workbook = new Workbook();
		WorkbookDesigner designer = new WorkbookDesigner();
		designer.setWorkbook(workbook);
		workbook.getWorksheets().get(0).getCells().get("A1").putValue("&=$VariableArray(HTML)");
		designer.setDataSource("VariableArray",
				new String[] { "Hello <b>World</b>", "Arabic", "Hindi", "Urdu", "French" });
		designer.process();
		workbook.save(dataDir + "UHProperty-out.xls");
	}
}
