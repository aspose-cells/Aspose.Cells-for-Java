package com.aspose.cells.examples.articles;

import com.aspose.cells.ActiveXControl;
import com.aspose.cells.ComboBoxActiveXControl;
import com.aspose.cells.ControlType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.Shape;

public class UpdateActiveXComboBoxControl {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UpdateActiveXComboBoxControl.class) + "articles/";
		// Create a workbook
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		// Access first shape from first worksheet
		Shape shape = wb.getWorksheets().get(0).getShapes().get(0);

		// Access ActiveX ComboBox Control and update its value
		if (shape.getActiveXControl() != null) {
			// Access Shape ActiveX Control
			ActiveXControl c = shape.getActiveXControl();

			// Check if ActiveX Control is ComboBox Control
			if (c.getType() == ControlType.COMBO_BOX) {
				// Type cast ActiveXControl into ComboBoxActiveXControl and
				// change its value
				ComboBoxActiveXControl comboBoxActiveX = (ComboBoxActiveXControl) c;
				comboBoxActiveX.setValue("This is combo box control.");
			}
		}

		// Save the workbook
		wb.save(dataDir + "UpdateActiveXComboBoxControl_out.xlsx");
	}
}
