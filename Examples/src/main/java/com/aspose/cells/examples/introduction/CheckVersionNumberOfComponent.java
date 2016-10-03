package com.aspose.cells.examples.introduction;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;

public class CheckVersionNumberOfComponent {
	public static void main(String[] args) throws Exception {
		try {
			// Instantiating a Workbook object
			Workbook workbook = new Workbook();
			System.out.println(CellsHelper.getVersion());
		}
		catch (Exception ee) {
			System.out.println(ee);
		}
	}
}