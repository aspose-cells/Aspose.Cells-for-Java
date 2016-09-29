package com.aspose.cells.examples.articles;

import java.io.FileOutputStream;

import com.aspose.cells.Cell;
import com.aspose.cells.DataBar;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class GenerateConditionalFormatting {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GenerateConditionalFormatting.class) + "articles/";
		// Create workbook object from source excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the cell which contains conditional formatting databar
		Cell cell = worksheet.getCells().get("C1");

		// Get the conditional formatting of the cell
		FormatConditionCollection fcc = cell.getFormatConditions();

		// Access the conditional formatting databar
		DataBar dbar = fcc.get(0).getDataBar();

		// Create image or print options
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setImageFormat(ImageFormat.getPng());

		// Get the image bytes of the databar
		byte[] imgBytes = dbar.toImage(cell, opts);

		// Write image bytes on the disk
		FileOutputStream out = new FileOutputStream(dataDir + "GCFormatting_out.png");
		out.write(imgBytes);
		out.close();

	}
}
