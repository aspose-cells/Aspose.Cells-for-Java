package com.aspose.cells.examples.articles;

import com.aspose.cells.Color;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.MsoPresetTextEffect;
import com.aspose.cells.Shape;
import com.aspose.cells.ShapeLockType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class LockWordArtWatermark {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(LockWordArtWatermark.class) + "articles/";

		// Instantiate a new Workbook
		Workbook workbook = new Workbook();

		// Get the first default sheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Add Watermark
		Shape wordart = sheet.getShapes().addTextEffect(MsoPresetTextEffect.TEXT_EFFECT_1, "CONFIDENTIAL",
				"Arial Black", 50, false, true, 18, 8, 1, 1, 130, 800);

		// Get the fill format of the word art
		FillFormat wordArtFormat = wordart.getFill();

		// Set the color
		wordArtFormat.setOneColorGradient(Color.getRed(), 0.2, GradientStyleType.HORIZONTAL, 2);

		// Set the transparency
		wordArtFormat.setTransparency(0.9);

		// Make the line invisible
		wordart.setHasLine(false);

		// Lock Shape Aspects
		wordart.setLocked(true);
		wordart.setLockedProperty(ShapeLockType.SELECTION, true);
		wordart.setLockedProperty(ShapeLockType.SHAPE_TYPE, true);
		wordart.setLockedProperty(ShapeLockType.MOVE, true);
		wordart.setLockedProperty(ShapeLockType.RESIZE, true);
		wordart.setLockedProperty(ShapeLockType.TEXT, true);

		// Save the file
		workbook.save(dataDir + "LkWordArtWatermark_out.xls");

	}
}
