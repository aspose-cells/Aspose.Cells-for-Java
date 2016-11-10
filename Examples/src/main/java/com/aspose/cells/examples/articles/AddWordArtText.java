package com.aspose.cells.examples.articles;

import com.aspose.cells.PresetWordArtStyle;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddWordArtText {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddWordArtText.class) + "articles/";
		// Create workbook object
		Workbook wb = new Workbook();

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Add Word Art Text with Built-in Styles
		ws.getShapes().addWordArt(PresetWordArtStyle.WORD_ART_STYLE_1, "Aspose File Format APIs", 00, 0, 0, 0, 100, 800);
		ws.getShapes().addWordArt(PresetWordArtStyle.WORD_ART_STYLE_2, "Aspose File Format APIs", 10, 0, 0, 0, 100, 800);
		ws.getShapes().addWordArt(PresetWordArtStyle.WORD_ART_STYLE_3, "Aspose File Format APIs", 20, 0, 0, 0, 100, 800);
		ws.getShapes().addWordArt(PresetWordArtStyle.WORD_ART_STYLE_4, "Aspose File Format APIs", 30, 0, 0, 0, 100, 800);
		ws.getShapes().addWordArt(PresetWordArtStyle.WORD_ART_STYLE_5, "Aspose File Format APIs", 40, 0, 0, 0, 100, 800);

		// Save the workbook in xlsx format
		wb.save(dataDir + "AddWordArtText_out.xlsx");
	}
}
