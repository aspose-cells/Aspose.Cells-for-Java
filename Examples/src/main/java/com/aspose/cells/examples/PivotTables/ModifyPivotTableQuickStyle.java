package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.Color;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Style;
import com.aspose.cells.TableStyle;
import com.aspose.cells.TableStyleElement;
import com.aspose.cells.TableStyleElementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ModifyPivotTableQuickStyle {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ModifyPivotTableQuickStyle.class) + "PivotTables/";
		// Open the template file containing the pivot table.
		Workbook wb = new Workbook(dataDir + "sample1.xlsx");
		// Add Pivot Table style
		Style style1 = wb.createStyle();
		com.aspose.cells.Font font1 = style1.getFont();
		font1.setColor(Color.getRed());
		Style style2 = wb.createStyle();
		com.aspose.cells.Font font2 = style2.getFont();
		font2.setColor(Color.getBlue());
		int i = wb.getWorksheets().getTableStyles().addPivotTableStyle("tt");
		// Get and Set the table style for different categories
		TableStyle ts = wb.getWorksheets().getTableStyles().get(i);
		int index = ts.getTableStyleElements().add(TableStyleElementType.FIRST_COLUMN);
		TableStyleElement e = ts.getTableStyleElements().get(index);
		e.setElementStyle(style1);
		index = ts.getTableStyleElements().add(TableStyleElementType.GRAND_TOTAL_ROW);
		e = ts.getTableStyleElements().get(index);
		e.setElementStyle(style2);

		// Set Pivot Table style name
		PivotTable pt = wb.getWorksheets().get(0).getPivotTables().get(0);
		pt.setPivotTableStyleName("tt");

		// Save the file.
		wb.save(dataDir + "ModifyPivotTableQuickStyle_out.xlsx");
	}
}
