package com.aspose.cells.examples.introduction;

import java.io.FileInputStream;

import com.aspose.cells.Cell;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.License;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.formulas.DirectCalculationFormula;

public class OpeningExistingFile {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(OpeningExistingFile.class) + "introduction/";
		// Creating a file input stream to reference the license file
		FileInputStream fstream = new FileInputStream("Aspose.Cells.lic");

		// Create a License object
		License license = new License();

		// Set the license of Aspose.Cells to avoid the evaluation limitations
		license.setLicense(fstream);

		// Instantiate a Workbook object that represents an Excel file
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Get the reference of "A1" cell from the cells of a worksheet
		Cell cell = workbook.getWorksheets().get(0).getCells().get("A1");

		// Set the "Hello World!" value into the "A1" cell
		cell.setValue("Hello World!");

		// Write the Excel file
		workbook.save(dataDir + "OpeningExistingFile_out.xls", FileFormatType.EXCEL_97_TO_2003);
		workbook.save(dataDir + "OpeningExistingFile_out.xlsx");
		workbook.save(dataDir + "OpeningExistingFile_out.ods");
	}
}
