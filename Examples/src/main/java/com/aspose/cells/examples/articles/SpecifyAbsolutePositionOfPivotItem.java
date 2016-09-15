package com.aspose.cells.examples.articles;

import com.aspose.cells.PivotField;
import com.aspose.cells.PivotFieldSubtotalType;
import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SpecifyAbsolutePositionOfPivotItem {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SpecifyAbsolutePositionOfPivotItem.class) + "articles/";
		Workbook wb = new Workbook(dataDir + "source.xlsx");
		Worksheet wsPivot = wb.getWorksheets().add("pvtNew Hardware");
		Worksheet wsData = wb.getWorksheets().get("New Hardware - Yearly");

		// Get the pivottables collection for the pivot sheet
		PivotTableCollection pivotTables = wsPivot.getPivotTables();

		// Add PivotTable to the worksheet
		int index = pivotTables.add("='New Hardware - Yearly'!A1:D621", "A3", "HWCounts_PivotTable");

		// Get the PivotTable object
		PivotTable pvtTable = pivotTables.get(index);

		// Add vendor row field
		pvtTable.addFieldToArea(PivotFieldType.ROW, "Vendor");

		// Add item row field
		pvtTable.addFieldToArea(PivotFieldType.ROW, "Item");

		// Add data field
		pvtTable.addFieldToArea(PivotFieldType.DATA, "2014");

		// Turn off the subtotals for the vendor row field
		PivotField pivotField = pvtTable.getRowFields().get("Vendor");
		pivotField.setSubtotals(PivotFieldSubtotalType.NONE, true);

		// Turn off grand total
		pvtTable.setColumnGrand(false);

		// Please call the PivotTable.RefreshData() and PivotTable.CalculateData(). Before using PivotItem.Position,
		// PivotItem.PositionInSameParentNode and PivotItem.Move(int count, bool isSameParent).
		pvtTable.refreshData();
		pvtTable.calculateData();

		pvtTable.getRowFields().get("Item").getPivotItems().get("4H12").setPositionInSameParentNode(0);
		pvtTable.getRowFields().get("Item").getPivotItems().get("DIF400").setPositionInSameParentNode(3);

		/*
		 * As a result of using PivotItem.PositionInSameParentNode,it will change the original sort sequence, // so when you use
		 * PivotItem.PositionInSameParentNode in another parent node,you need call the method named "CalculateData" again.
		 */
		pvtTable.calculateData();

		pvtTable.getRowFields().get("Item").getPivotItems().get("CA32").setPositionInSameParentNode(1);
		pvtTable.getRowFields().get("Item").getPivotItems().get("AAA3").setPositionInSameParentNode(2);

		// Save file
		wb.save(dataDir + "SAPOfPivotItem.xlsx");

	}
}
