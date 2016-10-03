package com.aspose.cells.examples.data;

import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SettingFontStyle {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getSharedDataDir(SettingFontStyle.class) + "data/";
		
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		FormatCondition fc = null;
		Style style = fc.getStyle();
		Font font = style.getFont();
		font.setItalic(true);
		font.setBold(true);
		font.setStrikeout(true);
		font.setUnderline(FontUnderlineType.DOUBLE);
		font.setColor(Color.getBlack());
		fc.setStyle(style);
		// Saving the Excel file
		workbook.save(dataDir + "SFontStyle_out.xls");
	}
}
