package com.aspose.cells.examples.SmartMarkers;

import java.util.ArrayList;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.examples.Utils;

public class UsingNestedObjects {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingNestedObjects.class) + "SmartMarkers/";
		Workbook workbook = new Workbook(dataDir + "TestSmartMarkers.xlsx");

		WorkbookDesigner designer = new WorkbookDesigner();
		designer.setWorkbook(workbook);

		ArrayList<Individual> list = new ArrayList<Individual>();
		list.add(new Individual("John", 23, new Wife("Jill", 20)));
		list.add(new Individual("Jack", 25, new Wife("Hilly", 21)));
		list.add(new Individual("James", 26, new Wife("Hally", 22)));
		list.add(new Individual("Baptist", 27, new Wife("Newly", 23)));

		designer.setDataSource("Individual", list);

		designer.process(false);

		workbook.save(dataDir + "UsingNestedObjects-out.xlsx");
	}

}
