package com.aspose.cells.examples.articles;

import com.aspose.cells.Color;
import com.aspose.cells.PresetShadowType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingTextEffectsShadowOfShapeOrTextbox {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingTextEffectsShadowOfShapeOrTextbox.class) + "articles/";

		// Create workbook object
		Workbook wb = new Workbook();

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Add text box with these dimensions
		int idx = ws.getTextBoxes().add(2, 2, 100, 400);
		TextBox tb = ws.getTextBoxes().get(idx);

		// Set the text of the textbox
		tb.setText("This text has the following settings.\n\nText Effects > Shadow > Offset Bottom");

		// Set all the text runs shadow to preset offset bottom
		for (int i = 0; i < tb.getTextBody().getCount(); i++) {
			tb.getTextBody().get(i).getShapeFont().getFillFormat().getShadowEffect()
					.setPresetType(PresetShadowType.OFFSET_BOTTOM);
		}

		// Set the font color and size of the textbox
		tb.getFont().setColor(Color.getRed());
		tb.getFont().setSize(16);

		// Save the output file
		wb.save(dataDir + "STESOfShapeOrTextbox_out.xlsx", SaveFormat.XLSX);

	}

}
