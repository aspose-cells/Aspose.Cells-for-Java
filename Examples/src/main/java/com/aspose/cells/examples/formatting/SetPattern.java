package com.aspose.cells.examples.formatting;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Color;
import com.aspose.cells.Style;
import com.aspose.cells.examples.Utils;

public class SetPattern {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getDataDir(SetBorder.class);
		Style style = fc.getStyle();
		style.setPattern(BackgroundType.REVERSE_DIAGONAL_STRIPE);
		style.setForegroundColor(Color.fromArgb(255, 255, 0));
		style.setBackgroundColor(Color.fromArgb(0, 255, 255));
		fc.setStyle(style);
	}
}
