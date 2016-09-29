package com.aspose.cells.examples.articles;

import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.TextParagraph;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CreateTextBoxhavingdifferentLineAlignment {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateTextBoxhavingdifferentLineAlignment.class) + "articles/";

		// Create a workbook.
		Workbook wb = new Workbook();

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Add text box inside the sheet.
		ws.getShapes().addShape(MsoDrawingType.TEXT_BOX, 2, 0, 2, 0, 80, 400);

		// Access first shape which is a text box and set is text.
		Shape shape = ws.getShapes().get(0);
		shape.setText(
				"Sign up for your free phone number.\nCall and text online for free.\nCall your friends and family.");

		// Acccess the first paragraph and set its horizontal alignment to left.
		TextParagraph p = shape.getTextBody().getTextParagraphs().get(0);
		p.setAlignmentType(TextAlignmentType.LEFT);

		// Acccess the second paragraph and set its horizontal alignment to center.
		p = shape.getTextBody().getTextParagraphs().get(1);
		p.setAlignmentType(TextAlignmentType.CENTER);

		// Acccess the third paragraph and set its horizontal alignment to right.
		p = shape.getTextBody().getTextParagraphs().get(2);
		p.setAlignmentType(TextAlignmentType.RIGHT);

		// Save the workbook in xlsx format.
		wb.save(dataDir + "CTBoxHDLineAlignment_out.xlsx", SaveFormat.XLSX);

	}

}
