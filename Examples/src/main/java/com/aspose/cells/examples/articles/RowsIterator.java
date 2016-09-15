package com.aspose.cells.examples.articles;

import java.util.Iterator;

import com.aspose.cells.Row;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RowsIterator {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(RowsIterator.class) + "articles/";
		// Load a file in an instance of Workbook
		Workbook book = new Workbook(dataDir + "sample.xlsx");

		// Get the iterator for RowCollection
		Iterator rowsIterator = book.getWorksheets().get(0).getCells().getRows().iterator();
		// Traverse rows in the collection
		while (rowsIterator.hasNext()) {
			Row row = (Row) rowsIterator.next();
			System.out.println(row.getIndex());
		}

	}
}
