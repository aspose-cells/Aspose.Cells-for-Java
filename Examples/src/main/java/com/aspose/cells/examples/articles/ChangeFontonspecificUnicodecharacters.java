package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ChangeFontonspecificUnicodecharacters {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeFontonspecificUnicodecharacters.class) + "articles/";

		// Create workbook object
		Workbook workbook = new Workbook();

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access cells
		Cell cell1 = worksheet.getCells().get("A1");
		Cell cell2 = worksheet.getCells().get("B1");

		// Set the styles of both cells to Times New Roman
		Style style = cell1.getStyle();
		style.getFont().setName("Times New Roman");
		cell1.setStyle(style);
		cell2.setStyle(style);

		// Put the values inside the cell
		cell1.putValue("Hello without Non-Breaking Hyphen");
		cell2.putValue("Hello" + (char) (8209) + " with Non-Breaking Hyphen");

		// Autofit the columns
		worksheet.autoFitColumns();

		// Save to Pdf without setting PdfSaveOptions.IsFontSubstitutionCharGranularity
		workbook.save(dataDir + "CFOnSUCharacters1_out.pdf");

		// Save to Pdf after setting PdfSaveOptions.IsFontSubstitutionCharGranularity to true
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setFontSubstitutionCharGranularity(true);
		workbook.save(dataDir + "CFOnSUCharacters2_out.pdf", opts);


	}
}
