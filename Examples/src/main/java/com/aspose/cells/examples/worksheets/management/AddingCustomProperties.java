package com.aspose.cells.examples.worksheets.management;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddingCustomProperties {
	public static void main(String[] args) throws Exception {
		// ExStart:AddingCustomProperties
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingCustomProperties.class) + "worksheets/";
		// Create workbook object
		Workbook workbook = new Workbook(FileFormatType.XLSX);

		// Add simple property without any type
		workbook.getContentTypeProperties().add("MK31", "Simple Data");

		// Add date time property with type
		workbook.getContentTypeProperties().add("MK32", "04-Mar-2015", "DateTime");

		// Save the workbook
		workbook.save(dataDir + "ACProperties-out.xlsx");
		// ExEnd:AddingCustomProperties
	}

}
