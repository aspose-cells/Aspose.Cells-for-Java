package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

// ExStart:1
class TextParser implements ICustomParser
{
	@Override
	public Object parseObject(String s) {
		return s;
	}

	@Override
	public String getFormat() {
		return "";
	}
}

class DateParser implements ICustomParser {
	@Override
	public Object parseObject(String s) {
		Date myDate = null;
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
		try {
			myDate = formatter.parse(s);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return myDate;
	}

	@Override
	public String getFormat() {
		return "dd/MM/yyyy";
	}
}

public class OpeningCSVFilesWithPreferredParser {

	//Source directory
	private static String sourceDir = Utils.Get_SourceDirectory();
	private static String outputDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// Initialize Text File's Load options
		TxtLoadOptions oTxtLoadOptions = new TxtLoadOptions(LoadFormat.CSV);

		// Specify the separatot character
		oTxtLoadOptions.setSeparator(',');

		// Specify the encoding scheme
		oTxtLoadOptions.setEncoding(Encoding.getUTF8());

		// Set the flag to true for converting datetime data
		oTxtLoadOptions.setConvertDateTimeData(true);

		// Set the preferred parsers
		oTxtLoadOptions.setPreferredParsers(new ICustomParser[] { new TextParser(), new DateParser() });

		// Initialize the workbook object by passing CSV file and text load options
		Workbook oExcelWorkBook = new Workbook(sourceDir + "samplePreferredParser.csv", oTxtLoadOptions);

		// Get the first cell
		Cell oCell = oExcelWorkBook.getWorksheets().get(0).getCells().get("A1");

		// Display type of value
		System.out.println("A1: " + getCellType(oCell.getType()) + " - " + oCell.getDisplayStringValue());

		// Get the second cell
		oCell = oExcelWorkBook.getWorksheets().get(0).getCells().get("B1");

		// Display type of value
		System.out.println("B1: " + getCellType(oCell.getType()) + " - " + oCell.getDisplayStringValue());

		// Save the workbook to disc
		oExcelWorkBook.save(outputDir + "outputsamplePreferredParser.xlsx");

		System.out.println("OpeningCSVFilesWithPreferredParser executed successfully.\r\n");
	}

	private static String getCellType(int type){
		if(type == CellValueType.IS_STRING){
			return "String";
		} else if(type == CellValueType.IS_NUMERIC){
			return "Numeric";
		} else if(type == CellValueType.IS_BOOL){
			return "Bool";
		} else if(type == CellValueType.IS_DATE_TIME){
			return "Date";
		} else if(type == CellValueType.IS_NULL){
			return "Null";
		} else if(type == CellValueType.IS_ERROR){
			return "Error";
		} else{
			return "Unknown";
		}
	}
	// ExEnd:1
}
