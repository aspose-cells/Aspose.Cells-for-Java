package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.TxtLoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertTexttoCols {

	public static void main(String[] args) throws Exception 
	{
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertTexttoCols.class) + "rows_cloumns/";
		
		//Create a workbook.
		Workbook wb = new Workbook();
		  
		//Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Add people name in column A. Fast name and Last name are separated by space.
		ws.getCells().get("A1").putValue("John Teal");
		ws.getCells().get("A2").putValue("Peter Graham");
		ws.getCells().get("A3").putValue("Brady Cortez");
		ws.getCells().get("A4").putValue("Mack Nick");
		ws.getCells().get("A5").putValue("Hsu Lee");
		  
		//Create text load options with space as separator.
		TxtLoadOptions opts = new TxtLoadOptions();
		opts.setSeparator(' ');
		  
		//Split the column A into two columns using TextToColumns() method.
		//Now column A will have first name and column B will have second name.
		ws.getCells().textToColumns(0, 0, 5, opts);
		  
		//Save the workbook in xlsx format.
		wb.save(dataDir + "outputTextToColumns.xlsx");
	}

}
