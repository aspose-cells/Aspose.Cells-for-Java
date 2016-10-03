package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableAutoFormatType;
import com.aspose.cells.PivotTableStyleType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetAutoFormatandPivotTableStyleTypes {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetAutoFormatandPivotTableStyleTypes.class) + "PivotTables/";
		// Load a template file
		Workbook workbook = new Workbook(dataDir + "PivotTable.xls");
		int pivotindex = 0;
		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(pivotindex);
		// Accessing the PivotTable
		PivotTable pivotTable = worksheet.getPivotTables().get(pivotindex);
		
		//Setting the PivotTable report is automatically formatted for Excel 2003 formats
		pivotTable.setAutoFormat(true);
		//Setting the PivotTable atuoformat type.
		pivotTable.setAutoFormatType(PivotTableAutoFormatType.CLASSIC);

		//Setting the PivotTable's Styles for Excel 2007/2010 formats e.g XLSX.
		pivotTable.setPivotTableStyleType(PivotTableStyleType.PIVOT_TABLE_STYLE_LIGHT_1);
		
	}
}
