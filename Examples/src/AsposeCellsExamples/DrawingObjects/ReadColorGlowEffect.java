package com.aspose.cells.examples.DrawingObjects;

import java.io.FileOutputStream;

import com.aspose.cells.*;
import com.aspose.cells.examples.*;

public class ReadColorGlowEffect {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReadColorGlowEffect.class) + "DrawingObjects/";

		//Read the source excel file
		Workbook wb = new Workbook(dataDir + "sourceGlowEffectColor.xlsx");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access the shape
		Shape sh = ws.getShapes().get(0);

		//Read the glow effect color and its various properties
		GlowEffect ge = sh.getGlow();
		CellsColor clr = ge.getColor();
		System.out.println("Color: " + clr.getColor());
		System.out.println("ColorIndex: " + clr.getColorIndex());
		System.out.println("IsShapeColor: " + clr.isShapeColor());
		System.out.println("Transparency: " + clr.getTransparency());
		System.out.println("Type: " + clr.getType());
		
	}
}
