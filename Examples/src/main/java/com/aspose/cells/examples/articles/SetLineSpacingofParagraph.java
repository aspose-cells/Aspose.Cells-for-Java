package com.aspose.cells.examples.articles;

import com.aspose.cells.LineSpaceSizeType;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.TextParagraph;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetLineSpacingofParagraph {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetLineSpacingofParagraph.class) + "articles/";
		// Create a workbook
		Workbook wb = new Workbook();

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Add text box inside the sheet
		ws.getShapes().addShape(MsoDrawingType.TEXT_BOX, 2, 0, 2, 0, 100, 200);

		// Access first shape which is a text box and set is text
		Shape shape = ws.getShapes().get(0);
		shape.setText("Sign up for your free phone number.\nCall and text online for free.");

		// Acccess the first paragraph
		TextParagraph p = shape.getTextBody().getTextParagraphs().get(1);

		// Set the line space
		p.setLineSpaceSizeType(LineSpaceSizeType.POINTS);
		p.setLineSpace(20);

		// Set the space after
		p.setSpaceAfterSizeType(LineSpaceSizeType.POINTS);
		p.setSpaceAfter(10);

		// Set the space before
		p.setSpaceBeforeSizeType(LineSpaceSizeType.POINTS);
		p.setSpaceBefore(10);

		// Save the workbook in xlsx format
		wb.save(dataDir + "SLSpacingofParagraph_out.xlsx", SaveFormat.XLSX);

	}

}
