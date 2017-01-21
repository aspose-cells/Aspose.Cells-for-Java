package com.aspose.cells.examples.CellsHelperClass;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SignificantDigits {

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(SignificantDigits.class) + "CellsHelperClass/";
		
		System.out.println(dataDir);
		
		//By default, Aspose.Cells stores 17 significant digits unlike
		//MS-Excel which stores only 15 significant digits
		CellsHelper.setSignificantDigits(15);

		//Create workbook
		Workbook workbook = new Workbook();

		//Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Access cell A1
		Cell c = worksheet.getCells().get("A1");

		//Put double value, only 15 significant digits as specified by
		//CellsHelper.SignificantDigits above will be stored in excel file just like MS-Excel does
		c.putValue(1234567890.123451711);

		//Save the workbook
		workbook.save(dataDir + "out_SignificantDigits.xlsx");
	}
}
