package com.aspose.cells.examples.formatting;

import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.Style;
import com.aspose.cells.examples.Utils;

public class SetBorder {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getSharedDataDir(SetBorder.class) + "formatting/";
		Style style = fc.getStyle();
		style.setBorder(BorderType.LEFT_BORDER, CellBorderType.DASHED, Color.fromArgb(0, 255, 255));
		style.setBorder(BorderType.TOP_BORDER, CellBorderType.DASHED, Color.fromArgb(0, 255, 255));
		style.setBorder(BorderType.RIGHT_BORDER, CellBorderType.DASHED, Color.fromArgb(0, 255, 255));
		style.setBorder(BorderType.RIGHT_BORDER, CellBorderType.DASHED, Color.fromArgb(255, 255, 0));
		fc.setStyle(style);
	}
}
