package AsposeCellsExamples.SmartMarkers;

import java.util.ArrayList;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.Worksheet;

import AsposeCellsExamples.Utils;

public class UsingNestedObjects {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingNestedObjects.class) + "SmartMarkers/";
		Workbook workbook = new Workbook();
		Worksheet worksheet = workbook.getWorksheets().get(0);

		worksheet.getCells().get("A1").putValue("Person Name");
		worksheet.getCells().get("A2").putValue("&=Individual.Name");

		worksheet.getCells().get("B1").putValue("Person Age");
		worksheet.getCells().get("B2").putValue("&=Individual.Age");

		worksheet.getCells().get("C1").putValue("Wife Name");
		worksheet.getCells().get("C2").putValue("&=Individual.Wife.Name");

		worksheet.getCells().get("D1").putValue("Wife Age");
		worksheet.getCells().get("D2").putValue("&=Individual.Wife.Age");

		WorkbookDesigner designer = new WorkbookDesigner();
		designer.setWorkbook(workbook);

		ArrayList<Individual> list = new ArrayList<Individual>();
		list.add(new Individual("John", 23, new Person("Jill", 20)));
		list.add(new Individual("Jack", 25, new Person("Hilly", 21)));
		list.add(new Individual("James", 26, new Person("Hally", 22)));
		list.add(new Individual("Baptist", 27, new Person("Newly", 23)));

		designer.setDataSource("Individual", list);

		designer.process(false);

		workbook.save(dataDir + "UsingNestedObjects-out.xlsx");
	}

}
