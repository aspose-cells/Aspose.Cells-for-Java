package com.aspose.cells.examples.articles;

import com.aspose.cells.VbaModule;
import com.aspose.cells.VbaModuleCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ModifyVBAorMacroCode {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ModifyVBAorMacroCode.class) + "articles/";
		// Create workbook object from source Excel file
		Workbook workbook = new Workbook(dataDir + "sample.xlsm");

		// Change the VBA Module Code
		VbaModuleCollection modules = workbook.getVbaProject().getModules();

		for (int i = 0; i < modules.getCount(); i++) {
			VbaModule module = modules.get(i);
			String code = module.getCodes();

			// Replace the original message with the modified message
			if (code.contains("This is test message.")) {
				code = code.replace("This is test message.", "This is Aspose.Cells message.");
				module.setCodes(code);
			}
		}

		// Save the output Excel file
		workbook.save(dataDir + "MVBAorMacroCode_out.xlsm");

	}
}
