package com.aspose.cells.examples.articles;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.ISmartMarkerCallBack;
import com.aspose.cells.Workbook;

public class SmartMarkerCallBack implements ISmartMarkerCallBack {
	Workbook workbook;

	SmartMarkerCallBack(Workbook workbook) {
		this.workbook = workbook;
	}

	@Override
	public void process(int sheetIndex, int rowIndex, int colIndex, String tableName, String columnName) {
		System.out.println("Processing Cell : " + workbook.getWorksheets().get(sheetIndex).getName() + "!"
				+ CellsHelper.cellIndexToName(rowIndex, colIndex));
		System.out.println("Processing Marker : " + tableName + "." + columnName);
	}
}