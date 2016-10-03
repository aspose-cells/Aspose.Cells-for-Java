package com.aspose.cells.examples.introduction;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class IndexToName {
	public static void main(String[] args) throws Exception {

		String cellname = CellsHelper.cellIndexToName(0, 0);
		System.out.println("Cell Name at [0, 0]: " + cellname);

		cellname = CellsHelper.cellIndexToName(4, 0);
		System.out.println("Cell Name at [4, 0]: " + cellname);

		cellname = CellsHelper.cellIndexToName(0, 4);
		System.out.println("Cell Name at [0, 4]: " + cellname);

		cellname = CellsHelper.cellIndexToName(2, 2);
		System.out.println("Cell Name at [2, 2]: " + cellname);

	}
}
