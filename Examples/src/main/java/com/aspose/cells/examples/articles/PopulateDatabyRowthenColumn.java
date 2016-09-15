package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class PopulateDatabyRowthenColumn {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PopulateDatabyRowthenColumn.class) + "articles/";
		Workbook workbook = new Workbook();
		Cells cells = workbook.getWorksheets().get(0).getCells();
		cells.get("A1").setValue("data1");
		cells.get("B1").setValue("data2");
		cells.get("A2").setValue("data3");
		cells.get("B2").setValue("data4");

	}
}
