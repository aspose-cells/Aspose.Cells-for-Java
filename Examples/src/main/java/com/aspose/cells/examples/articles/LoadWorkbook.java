package com.aspose.cells.examples.articles;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class LoadWorkbook {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(LoadWorkbook.class) + "articles/";

		// Create a sample workbook and add some data inside the first worksheet
		Workbook workbook = new Workbook();
		Worksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.getCells().get("P30").putValue("This is sample data.");

		// Save the workbook in memory stream
		ByteArrayOutputStream baout = new ByteArrayOutputStream();
		workbook.save(baout, SaveFormat.XLSX);

		// Get bytes and create byte array input stream
		byte[] bts = baout.toByteArray();
		ByteArrayInputStream bain = new ByteArrayInputStream(bts);

		// Now load the workbook from memory stream with A5 paper size
		LoadOptions opts = new LoadOptions(LoadFormat.XLSX);
		opts.setPaperSize(PaperSizeType.PAPER_A_5);
		workbook = new Workbook(bain, opts);

		// Save the workbook in pdf format
		workbook.save(dataDir + "output-a5.pdf");

		// Now load the workbook again from memory stream with A3 paper size
		opts = new LoadOptions(LoadFormat.XLSX);
		opts.setPaperSize(PaperSizeType.PAPER_A_3);
		workbook = new Workbook(bain, opts);

		// Save the workbook in pdf format
		workbook.save(dataDir + "LWorkbook_out.pdf");

	}

}
