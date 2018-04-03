package AsposeCellsExamples.Data;

import com.aspose.cells.*;

public class PreserveSingleQuotePrefixOfCellValueOrRange { 
	
	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create workbook
		Workbook wb = new Workbook();

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access cell A1
		Cell cell = ws.getCells().get("A1");

		//Put some text in cell, it does not have Single Quote at the beginning
		cell.putValue("Text");

		//Access style of cell A1
		Style st = cell.getStyle();

		//Print the value of Style.QuotePrefix of cell A1
		System.out.println("Quote Prefix of Cell A1: " + st.getQuotePrefix());

		//Put some text in cell, it has Single Quote at the beginning
		cell.putValue("'Text");

		//Access style of cell A1
		st = cell.getStyle();

		//Print the value of Style.QuotePrefix of cell A1
		System.out.println("Quote Prefix of Cell A1: " + st.getQuotePrefix());

		//Print information about StyleFlag.QuotePrefix property
		System.out.println();
		System.out.println("When StyleFlag.QuotePrefix is False, it means, do not update the value of Cell.Style.QuotePrefix.");
		System.out.println("Similarly, when StyleFlag.QuotePrefix is True, it means, update the value of Cell.Style.QuotePrefix.");
		System.out.println();

		//Create an empty style
		st = wb.createStyle();

		//Create style flag - set StyleFlag.QuotePrefix as false
		//It means, we do not want to update the Style.QuotePrefix property of cell A1's style.
		StyleFlag flag = new StyleFlag();
		flag.setQuotePrefix(false);

		//Create a range consisting of single cell A1
		Range rng = ws.getCells().createRange("A1");

		//Apply the style to the range
		rng.applyStyle(st, flag);

		//Access the style of cell A1
		st = cell.getStyle();

		//Print the value of Style.QuotePrefix of cell A1
		//It will print True, because we have not updated the Style.QuotePrefix property of cell A1's style.
		System.out.println("Quote Prefix of Cell A1: " + st.getQuotePrefix());

		//Create an empty style
		st = wb.createStyle();

		//Create style flag - set StyleFlag.QuotePrefix as true
		//It means, we want to update the Style.QuotePrefix property of cell A1's style.
		flag = new StyleFlag();
		flag.setQuotePrefix(true);

		//Apply the style to the range
		rng.applyStyle(st, flag);

		//Access the style of cell A1
		st = cell.getStyle();

		//Print the value of Style.QuotePrefix of cell A1
		//It will print False, because we have updated the Style.QuotePrefix property of cell A1's style.
		System.out.println("Quote Prefix of Cell A1: " + st.getQuotePrefix());

		// Print the message
		System.out.println("PreserveSingleQuotePrefixOfCellValueOrRange executed successfully.");
	}
}
