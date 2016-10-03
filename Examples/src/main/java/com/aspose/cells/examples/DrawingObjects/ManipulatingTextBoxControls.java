package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ManipulatingTextBoxControls {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ManipulatingTextBoxControls.class) + "DrawingObjects/";

		// Instantiate a new Workbook.
		Workbook workbook = new Workbook(dataDir + "tempBook1ole0.Xls");

		// Get the first worksheet in the book.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Get the first textbox object.
		com.aspose.cells.TextBox textbox0 = worksheet.getTextBoxes().get(0);

		// Obtain the text in the first textbox.
		String text0 = textbox0.getText();
		System.out.println(text0);

		// Get the second textbox object.
		com.aspose.cells.TextBox textbox1 = worksheet.getTextBoxes().get(1);

		// Obtain the text in the second textbox.
		String text1 = textbox1.getText();

		// Change the text of the second textbox.
		textbox1.setText("This is an alternative text");

		// Save the excel file.
		workbook.save(dataDir + "ManipulatingTextBoxControls_out.xls");
	}
}
