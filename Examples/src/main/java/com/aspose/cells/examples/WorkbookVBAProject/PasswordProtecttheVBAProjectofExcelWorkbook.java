package com.aspose.cells.examples.WorkbookVBAProject;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class PasswordProtecttheVBAProjectofExcelWorkbook {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PasswordProtecttheVBAProjectofExcelWorkbook.class) + "WorkbookVBAProject/";

		// Load your source Excel file.
		Workbook wb = new Workbook(dataDir + "samplePasswordProtectVBAProject.xlsm");

		// Access the VBA project of the workbook.
		VbaProject vbaProject = wb.getVbaProject();

		// Lock the VBA project for viewing with password.
		vbaProject.protect(true, "11");

		// Save the output Excel file
		wb.save(dataDir + "outputPasswordProtectVBAProject.xlsm");

		// Print message
		System.out.println("PasswordProtecttheVBAProjectofExcelWorkbook Done Successfully");
	}
}
