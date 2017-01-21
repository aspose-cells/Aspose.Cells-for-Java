package com.aspose.cells.examples.worksheets;

import com.aspose.cells.*;

public class GetPaperWidthHeight {
	public static void main(String[] args) throws Exception {

		//Create workbook
		Workbook wb = new Workbook();

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Set paper size to A2 and print paper width and height in inches
		ws.getPageSetup().setPaperSize(PaperSizeType.PAPER_A_2);
		System.out.println("PaperA2: " + ws.getPageSetup().getPaperWidth() + "x" + ws.getPageSetup().getPaperHeight());

		//Set paper size to A3 and print paper width and height in inches
		ws.getPageSetup().setPaperSize(PaperSizeType.PAPER_A_3);
		System.out.println("PaperA3: " + ws.getPageSetup().getPaperWidth() + "x" + ws.getPageSetup().getPaperHeight());

		//Set paper size to A4 and print paper width and height in inches
		ws.getPageSetup().setPaperSize(PaperSizeType.PAPER_A_4);
		System.out.println("PaperA4: " + ws.getPageSetup().getPaperWidth() + "x" + ws.getPageSetup().getPaperHeight());

		//Set paper size to Letter and print paper width and height in inches
		ws.getPageSetup().setPaperSize(PaperSizeType.PAPER_LETTER);
		System.out.println("PaperLetter: " + ws.getPageSetup().getPaperWidth() + "x" + ws.getPageSetup().getPaperHeight());
	}
}
