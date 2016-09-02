package com.aspose.cells.examples.formatting;

import com.aspose.cells.Font;
import com.aspose.cells.Style;

public class SettingStrikeOutEffect {
	public static void main(String[] args) throws Exception {
		// Setting the strike out effect on the font
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setStrikeout(true);
		cell.setStyle(style);
	}
}
