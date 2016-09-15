package com.aspose.cells.examples.articles;

import java.util.Iterator;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CellsIterator {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(CellsIterator.class) + "articles/";
		// Load a file in an instance of Workbook
		Workbook book = new Workbook(dataDir + "sample.xlsx");

		// Get the iterator from Cells collection
		Iterator cellIterator = book.getWorksheets().get(0).getCells().iterator();
		// Traverse cells in the collection
		while (cellIterator.hasNext()) {
			Cell cell = (Cell) cellIterator.next();
			;
			System.out.println(cell.getName() + " " + cell.getValue());
		}

		// Get iterator from an object of Row
		Iterator rowIterator = book.getWorksheets().get(0).getCells().getRows().get(0).iterator();
		// Traverse cells in the given row
		while (rowIterator.hasNext()) {
			Cell cell = (Cell) rowIterator.next();
			System.out.println(cell.getName() + " " + cell.getValue());
		}

		// Get iterator from an object of Range
		Iterator rangeIterator = book.getWorksheets().get(0).getCells().createRange("A1:B10").iterator();
		// Traverse cells in the range
		while (rangeIterator.hasNext()) {
			Cell cell = (Cell) rangeIterator.next();
			System.out.println(cell.getName() + " " + cell.getValue());
		}

	}
}
