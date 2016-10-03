package com.aspose.cells.examples.worksheets;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AdvancedProtectionSettingsUsingAsposeCells {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AdvancedProtectionSettingsUsingAsposeCells.class) + "worksheets/";

		// Instantiating a Workbook object by excel file path
		Workbook excel = new Workbook(dataDir + "book1.xlsx");

		WorksheetCollection worksheets = excel.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		Protection protection = worksheet.getProtection();

		// Restricting users to delete columns of the worksheet
		protection.setAllowDeletingColumn(false);

		// Restricting users to delete row of the worksheet
		protection.setAllowDeletingRow(false);

		// Restricting users to edit contents of the worksheet
		protection.setAllowEditingContent(false);

		// Restricting users to edit objects of the worksheet
		protection.setAllowEditingObject(false);

		// Restricting users to edit scenarios of the worksheet
		protection.setAllowEditingScenario(false);

		// Restricting users to filter
		protection.setAllowFiltering(false);

		// Allowing users to format cells of the worksheet
		protection.setAllowFormattingCell(true);

		// Allowing users to format rows of the worksheet
		protection.setAllowFormattingRow(true);

		// Allowing users to insert columns in the worksheet
		protection.setAllowInsertingColumn(true);

		// Allowing users to insert hyperlinks in the worksheet
		protection.setAllowInsertingHyperlink(true);

		// Allowing users to insert rows in the worksheet
		protection.setAllowInsertingRow(true);

		// Allowing users to select locked cells of the worksheet
		protection.setAllowSelectingLockedCell(true);

		// Allowing users to select unlocked cells of the worksheet
		protection.setAllowSelectingUnlockedCell(true);

		// Allowing users to sort
		protection.setAllowSorting(true);

		// Allowing users to use pivot tables in the worksheet
		protection.setAllowUsingPivotTable(true);

		// Saving the modified Excel file Excel XP format
		excel.save(dataDir + "APSettingsUsingAsposeCells_out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Worksheet protected successfully.");

	}
}
