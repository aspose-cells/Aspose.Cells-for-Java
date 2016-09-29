package com.aspose.cells.examples.articles;

import com.aspose.cells.BuiltinStyleType;
import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class UsingBuiltinStyles {

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(UsingBuiltinStyles.class) + "articles/";
		String output1Path = dataDir + "UsingBuiltinStyles_out.xlsx";
		String output2Path = dataDir + "UsingBuiltinStyles_out.ods";

		Workbook workbook = new Workbook();
		Style style = workbook.createBuiltinStyle(BuiltinStyleType.TITLE);

		Cell cell = workbook.getWorksheets().get(0).getCells().get("A1");
		cell.putValue("Aspose");
		cell.setStyle(style);

		workbook.getWorksheets().get(0).autoFitColumn(0);
		workbook.getWorksheets().get(0).autoFitRow(0);

		workbook.save(output1Path);
		System.out.println("File saved " + output1Path);
		workbook.save(output2Path);
		System.out.println("File saved " + output2Path);

	}
}
