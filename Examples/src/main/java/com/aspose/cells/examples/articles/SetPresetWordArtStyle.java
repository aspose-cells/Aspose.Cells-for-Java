package com.aspose.cells.examples.articles;

import java.util.ArrayList;

import com.aspose.cells.FontSetting;
import com.aspose.cells.PresetWordArtStyle;
import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetPresetWordArtStyle {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetPresetWordArtStyle.class) + "articles/";

		//Create workbook object
		Workbook wb = new Workbook();

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Create a textbox with some text
		int idx = ws.getTextBoxes().add(0, 0, 100, 700);
		TextBox tb = ws.getTextBoxes().get(idx);
		tb.setText("Aspose File Format APIs");
		tb.getFont().setSize(44);

		//Sets preset WordArt style to the text of the shape.
		ArrayList<FontSetting> aList = tb.getCharacters();
		FontSetting fntSetting = aList.get(0);
		
		fntSetting.setWordArtStyle(PresetWordArtStyle.WORD_ART_STYLE_3);

		//Save the workbook in xlsx format
		wb.save(dataDir + "SetPresetWordArtStyle_out.xlsx");
	}
}
