package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableAutoFormatType;
import com.aspose.cells.PivotTableStyleType;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.introduction.OpeningExistingFile;

public class SetAutoFormatandPivotTableStyleTypes {
	public static void main(String[] args) throws Exception {
		
		//Setting the PivotTable report is automatically formatted for Excel 2003 formats
		pivotTable.setAutoFormat(true);
		//Setting the PivotTable atuoformat type.
		pivotTable.setAutoFormatType(PivotTableAutoFormatType.CLASSIC);

		//Setting the PivotTable's Styles for Excel 2007/2010 formats e.g XLSX.
		pivotTable.setPivotTableStyleType(PivotTableStyleType.PIVOT_TABLE_STYLE_LIGHT_1);
		
	}
}
