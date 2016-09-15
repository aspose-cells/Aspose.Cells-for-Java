package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.MemorySetting;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class WritingLargeExcelFiles {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(WritingLargeExcelFiles.class) + "articles/";
		// Instantiate a new Workbook
		Workbook wb = new Workbook();
		// Set the memory preferences
		// Note: This setting cannot take effect for the existing worksheets that are created before using the below line of code
		wb.getSettings().setMemorySetting(MemorySetting.MEMORY_PREFERENCE);

		/*
		 * Note: The memory settings also would not work for the default sheet i.e., "Sheet1" etc. automatically created by the
		 * Workbook. To change the memory setting of existing sheets, please change memory setting for them manually:
		 */
		Cells cells = wb.getWorksheets().get(0).getCells();
		cells.setMemorySetting(MemorySetting.MEMORY_PREFERENCE);
		// Input large dataset into the cells of the worksheet.Your code goes here.

		// Get cells of the newly created Worksheet "Sheet2" whose memory setting is same with the one defined in
		// WorkbookSettings:
		cells = wb.getWorksheets().add("Sheet2").getCells();

	}
}
