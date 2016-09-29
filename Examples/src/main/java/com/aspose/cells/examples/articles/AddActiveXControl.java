package com.aspose.cells.examples.articles;

import com.aspose.cells.ActiveXControl;
import com.aspose.cells.ControlType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddActiveXControl {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddActiveXControl.class) + "articles/";
		// Create workbook object
		Workbook wb = new Workbook();

		// Access first worksheet
		Worksheet sheet = wb.getWorksheets().get(0);

		// Add Toggle Button ActiveX Control inside the Shape Collection
		Shape s = sheet.getShapes().addActiveXControl(ControlType.TOGGLE_BUTTON, 4, 0, 4, 0, 100, 30);

		// Access the ActiveX control object and set its linked cell property
		ActiveXControl c = s.getActiveXControl();
		c.setLinkedCell("A1");

		// Save the worbook in xlsx format
		wb.save(dataDir + "AAXControl_out.xlsx", SaveFormat.XLSX);

	}
}
