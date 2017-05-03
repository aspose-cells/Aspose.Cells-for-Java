package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SpecifyingDBNumCustomPatternFormatting {

	public static void main(String[] args) throws Exception {
		
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SpecifyingDBNumCustomPatternFormatting.class) + "data/";

		//Create a workbook.
		Workbook wb = new Workbook();
		  
		//Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Access cell A1 and put value 123.
		Cell cell = ws.getCells().get("A1");
		cell.putValue(123);
		  
		//Access cell style.
		Style st = cell.getStyle();
		  
		//Specifying DBNum custom pattern formatting.
		st.setCustom("[DBNum2][$-804]General");
		  
		//Set the cell style.
		cell.setStyle(st);
		  
		//Set the first column width.
		ws.getCells().setColumnWidth(0, 30);
		  
		//Save the workbook in output pdf format.
		wb.save(dataDir + "outputDBNumCustomFormatting.pdf", SaveFormat.PDF);

		// Print message
		System.out.println("SpecifyingDBNumCustomPatternFormatting Done Successfully");

	}
}
