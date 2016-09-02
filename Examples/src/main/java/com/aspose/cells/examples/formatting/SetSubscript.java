package com.aspose.cells.examples.formatting;

import com.aspose.cells.Font;
import com.aspose.cells.Style;

public class SetSubscript {
	public static void main(String[] args) throws Exception {
		// Setting subscript effect
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setSubscript(true);
		cell.setStyle(style);
	}
}
