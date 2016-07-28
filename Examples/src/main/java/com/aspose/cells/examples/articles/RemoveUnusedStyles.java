package com.aspose.cells.examples.articles;

import com.aspose.cells.VbaProject;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RemoveUnusedStyles {

	public static void main(String[] args) throws Exception {
		// ExStart:RemoveUnusedStyles
		String dataDir = Utils.getDataDir(RemoveUnusedStyles.class);
		String inputPath = dataDir + "Styles.xlsx";
		String outputPath = dataDir + "Output.xlsx";

		Workbook workbook = new Workbook(inputPath);

		workbook.removeUnusedStyles();

		workbook.save(outputPath);
		System.out.println("File saved " + outputPath);
		// ExEnd:RemoveUnusedStyles
	}
}
