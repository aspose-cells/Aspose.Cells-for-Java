package com.aspose.cells.examples.LoadingSavingConvertingAndManaging;



import com.aspose.cells.SaveFormat;
import com.aspose.cells.TxtSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class TrimBlankRowsAndColsWhileExportingSpreadsheetsToCSVFormat {

	public static void main(String[] args) throws Exception{
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(TrimBlankRowsAndColsWhileExportingSpreadsheetsToCSVFormat.class) + "LoadingSavingConvertingAndManaging/";
		//Load source worbook
		Workbook wb = new Workbook(dataDir + "sampleTrimBlankColumns.xlsx");
		 
		//Save in csv format
		wb.save(dataDir + "outputWithoutTrimBlankColumns.csv", SaveFormat.CSV);
		 
		//Now save again with TrimLeadingBlankRowAndColumn as true
		TxtSaveOptions opts = new TxtSaveOptions();
		opts.setTrimLeadingBlankRowAndColumn(true);
		 
		//Save in csv format
		wb.save(dataDir + "outputTrimBlankColumns.csv", opts);
	}

}
