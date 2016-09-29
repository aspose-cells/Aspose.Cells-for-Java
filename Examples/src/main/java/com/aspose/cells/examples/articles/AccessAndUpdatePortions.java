package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.FontSetting;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.AdvancedTopics.SmartMarkers.UsingHTMLProperty;

public class AccessAndUpdatePortions {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AccessAndUpdatePortions.class) + "articles/";

		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		Cell cell = worksheet.getCells().get("A1");

		System.out.println("Before updating the font settings....");

		FontSetting[] fnts = cell.getCharacters();

		for (int i = 0; i < fnts.length; i++) {
			System.out.println(fnts[i].getFont().getName());
		}

		// Modify the first FontSetting Font Name
		fnts[0].getFont().setName("Arial");

		// And update it using SetCharacters() method
		cell.setCharacters(fnts);

		System.out.println();

		System.out.println("After updating the font settings....");

		fnts = cell.getCharacters();

		for (int i = 0; i < fnts.length; i++) {
			System.out.println(fnts[i].getFont().getName());
		}

		// Save workbook
		workbook.save(dataDir + "AAUPortions_out.xlsx");

	}
}
