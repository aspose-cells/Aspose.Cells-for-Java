package com.aspose.cells.examples.articles;

import com.aspose.cells.Picture;
import com.aspose.cells.SignatureLine;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CreateSignatureLine {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateSignatureLine.class) + "articles/";
		// Create workbook object
		Workbook workbook = new Workbook();

		// Insert picture of your choice
		int index = workbook.getWorksheets().get(0).getPictures().add(0, 0, "signature.jpg");

		// Access picture and add signature line inside it
		Picture pic = workbook.getWorksheets().get(0).getPictures().get(index);

		// Create signature line object
		SignatureLine s = new SignatureLine();
		s.setSigner("Simon Zhao");
		s.setTitle("Development Lead");
		s.setEmail("Simon.Zhao@aspose.com");

		// Assign the signature line object to Picture.SignatureLine property
		pic.setSignatureLine(s);

		// Save the workbook
		workbook.save(dataDir + "CSignatureLine_out.xlsx");

	}
}
