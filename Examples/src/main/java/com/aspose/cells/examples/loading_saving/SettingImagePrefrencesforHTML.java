package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SettingImagePrefrencesforHTML {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingImagePrefrencesforHTML.class) + "loading_saving/";

		// Instantiate a Workbook object by excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");
		
		System.out.println("Set PDF Creation Time successfully.");

	}
}
