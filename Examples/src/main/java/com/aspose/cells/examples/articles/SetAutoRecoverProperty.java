package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SetAutoRecoverProperty {
	public static void main(String[] args) throws Exception {
		// ExStart:SetAutoRecoverProperty
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SetAutoRecoverProperty.class);
		// Create workbook object
		Workbook workbook = new Workbook();

		// Read AutoRecover property
		System.out.println("AutoRecover: " + workbook.getSettings().getAutoRecover());

		// Set AutoRecover property to false
		workbook.getSettings().setAutoRecover(false);

		// Save the workbook
		workbook.save("output.xlsx");

		// Read the saved workbook again
		workbook = new Workbook("output.xlsx");

		// Read AutoRecover property
		System.out.println("AutoRecover: " + workbook.getSettings().getAutoRecover());
		// ExEnd:SetAutoRecoverProperty
	}
}
