package com.aspose.cells.examples.articles;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.VbaModule;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddVBAModuleAndCode {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddVBAModuleAndCode.class) + "articles/";
		// Create new workbook
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add VBA Module
		int idx = workbook.getVbaProject().getModules().add(worksheet);

		// Access the VBA Module, set its name and codes
		VbaModule module = workbook.getVbaProject().getModules().get(idx);
		module.setName("TestModule");

		module.setCodes("Sub ShowMessage()" + "\r\n" + "    MsgBox \"Welcome to Aspose!\"" + "\r\n" + "End Sub");

		// Save the workbook
		workbook.save(dataDir + "AVBAMAndCode_out.xlsm", SaveFormat.XLSM);


	}
}
