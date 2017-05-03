package com.aspose.cells.examples.WorkbookVBAProject;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class CheckifVBAProjectisProtectedandLockedforViewing {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CheckifVBAProjectisProtectedandLockedforViewing.class) + "WorkbookVBAProject/";

		// Load your source Excel file.
		Workbook wb = new Workbook(dataDir + "sampleCheckifVBAProjectisProtected.xlsm");

		// Access the VBA project of the workbook.
		VbaProject vbaProject = wb.getVbaProject();

		// Whether "Lock project for viewing" is true or not.
		System.out.println("Is VBA Project Locked for Viewing: " + vbaProject.getIslockedForViewing());

		// Print message
		System.out.println("CheckifVBAProjectisProtectedandLockedforViewing Done Successfully");
	}
}
