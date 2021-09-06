package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class SavingFiletoSomeLocation {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingFiletoSomeLocation.class) + "LoadingSavingConvertingAndManaging/";

		String filePath = dataDir + "Book1.xls";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(filePath);

		// Save in Excel 97 – 2003 format
		workbook.save(dataDir + "SFTSomeLocation_out.xls");
		// OR
		// workbook.save(dataDir + ".output..xls", new
		// XlsSaveOptions(SaveFormat.Excel97To2003));

		// Save in Excel2007 xlsx format
		workbook.save(dataDir + "SFTSomeLocation_out.xlsx", SaveFormat.XLSX);

		// Save in Excel2007 xlsb format
		workbook.save(dataDir + "SFTSomeLocation_out.xlsb", SaveFormat.XLSB);

		// Save in ODS format
		workbook.save(dataDir + "SFTSomeLocation_out.ods", SaveFormat.ODS);

		// Save in Pdf format
		workbook.save(dataDir + "SFTSomeLocation_out.pdf", SaveFormat.PDF);

		// Save in Html format
		workbook.save(dataDir + "SFTSomeLocation_out.html", SaveFormat.HTML);

		// Save in SpreadsheetML format
		workbook.save(dataDir + "SFTSomeLocation_out.xml", SaveFormat.SPREADSHEET_ML);

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
