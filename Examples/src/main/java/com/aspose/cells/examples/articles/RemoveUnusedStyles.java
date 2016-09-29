package com.aspose.cells.examples.articles;

import com.aspose.cells.VbaProject;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RemoveUnusedStyles {

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(RemoveUnusedStyles.class) + "articles/";
		String inputPath = dataDir + "Styles.xlsx";
		String outputPath = dataDir + "RemoveUnusedStyles_out.xlsx";

		Workbook workbook = new Workbook(inputPath);

		workbook.removeUnusedStyles();

		workbook.save(outputPath);
		System.out.println("File saved " + outputPath);

	}
}
