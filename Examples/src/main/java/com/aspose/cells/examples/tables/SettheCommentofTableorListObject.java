package com.aspose.cells.examples.tables;

import com.aspose.cells.ListObject;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettheCommentofTableorListObject {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettheCommentofTableorListObject.class) + "tables/";

		// Open the template file.
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first list object or table.
		ListObject lstObj = worksheet.getListObjects().get(0);

		// Set the comment of the list object
		lstObj.setComment("This is Aspose.Cells comment.");

		// Save the workbook in xlsx format
		workbook.save(dataDir + "STheCofTOrListObject_out.xlsx", SaveFormat.XLSX);

	}

}
