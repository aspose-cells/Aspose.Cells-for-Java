package com.aspose.cells.examples.introduction;

import com.aspose.cells.CellsHelper;

public class NameToIndex {

	public static void main(String[] args) throws Exception {

		int[] cellIndices = CellsHelper.cellNameToIndex("C6");
		System.out.println("Row Index of Cell C6: " + cellIndices[0]);
		System.out.println("Column Index of Cell C6: " + cellIndices[1]);

	}
}
