package com.aspose.cells.examples.articles;

import java.io.ByteArrayInputStream;

import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class CopyPicturefromOneWorksheetToAnother {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyPicturefromOneWorksheetToAnother.class) + "articles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Shapes.xls");

		WorksheetCollection ws = workbook.getWorksheets();
		Worksheet sheet1 = ws.get("Picture");
		Worksheet sheet2 = ws.get("Result");

		// get the Picture from first worksheet
		Picture pic = sheet1.getPictures().get(0);

		ByteArrayInputStream bis = new ByteArrayInputStream(pic.getData());

		// Copy the Picture to Second Worksheet
		sheet2.getPictures().add(pic.getUpperLeftRow(), pic.getUpperLeftColumn(), pic.getWidthScale(),
				pic.getHeightScale(), bis);

		// Save the workbook
		workbook.save(dataDir + "CPfOneWToAnother_out.xls");

	}
}
