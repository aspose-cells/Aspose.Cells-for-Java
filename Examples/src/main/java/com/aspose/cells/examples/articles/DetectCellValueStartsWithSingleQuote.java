package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class DetectCellValueStartsWithSingleQuote {

	public static void main(String[] args) {

		// Create an instance of workbook
		Workbook workbook = new Workbook();

		// Access first worksheet from the collection
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access cells A1 and A2
		Cell a1 = worksheet.getCells().get("A1");
		Cell a2 = worksheet.getCells().get("A2");

		// Add simple text to cell A1 and text with quote prefix to cell A2
		a1.putValue("sample");
		a2.putValue("'sample");

		// Print their string values, A1 and A2 both are same
		System.out.println("String value of A1: " + a1.getStringValue());
		System.out.println("String value of A2: " + a2.getStringValue());

		// Access styles of cells A1 and A2
		Style s1 = a1.getStyle();
		Style s2 = a2.getStyle();

		System.out.println();

		// Check if A1 and A2 has a quote prefix
		System.out.println("A1 has a quote prefix: " + s1.getQuotePrefix());
		System.out.println("A2 has a quote prefix: " + s2.getQuotePrefix());

	}

}
