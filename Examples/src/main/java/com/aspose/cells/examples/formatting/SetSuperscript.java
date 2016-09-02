package com.aspose.cells.examples.formatting;

import com.aspose.cells.Font;
import com.aspose.cells.Style;

public class SetSuperscript {
	public static void main(String[] args) throws Exception {
		// Setting superscript effect
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setSuperscript(true);
		cell.setStyle(style);
	}
}
