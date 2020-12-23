package AsposeCellsExamples.SmartMarkers;

import java.util.ArrayList;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Color;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.Worksheet;

public class UsingGenericList {

	public static void main(String[] args) throws Exception {
		//ExStart: 1
		// Create a designer workbook
		Workbook workbook = new Workbook();
		Worksheet worksheet = workbook.getWorksheets().get(0);

		worksheet.getCells().get("A1").putValue("Teacher Name");
		worksheet.getCells().get("A2").putValue("&=Teacher.Name");

		worksheet.getCells().get("B1").putValue("Teacher Age");
		worksheet.getCells().get("B2").putValue("&=Teacher.Age");

		worksheet.getCells().get("C1").putValue("Student Name");
		worksheet.getCells().get("C2").putValue("&=Teacher.Students.Name");

		worksheet.getCells().get("D1").putValue("Student Age");
		worksheet.getCells().get("D2").putValue("&=Teacher.Students.Age");

		// Apply Style to A1:D1
		Range range = worksheet.getCells().createRange("A1:D1");
		Style style = workbook.createStyle();
		style.getFont().setBold(true);
		style.setForegroundColor(Color.getYellow());
		style.setPattern(BackgroundType.SOLID);
		StyleFlag flag = new StyleFlag();
		flag.setAll(true);
		range.applyStyle(style, flag);

		// Initialize WorkbookDesigner object
		WorkbookDesigner designer = new WorkbookDesigner();

		// Load the template file
		designer.setWorkbook(workbook);

		ArrayList<Teacher> list = new ArrayList<>();

		// Create the relevant student objects for the Teacher object
		ArrayList<Person> students = new ArrayList<>();
		students.add(new Person("Chen Zhao", 14));
		students.add(new Person("Jamima Winfrey", 18));
		students.add(new Person("Reham Smith", 15));

		// Create a Teacher object
		Teacher h1 = new Teacher("Mark John", 30, students);

		// Create the relevant student objects for the Teacher object
		students = new ArrayList<>();
		students.add(new Person("Karishma Jathool", 16));
		students.add(new Person("Angela Rose", 13));
		students.add(new Person("Hina Khanna", 15));

		// Create a Teacher object
		Teacher h2 = new Teacher("Masood Shankar", 40, students);

		// Add the objects to the list
		list.add(h1);
		list.add(h2);

		// Specify the DataSource
		designer.setDataSource("Teacher", list);

		// Process the markers
		designer.process();

		// Autofit columns
		worksheet.autoFitColumns();

		// Save the Excel file.
		designer.getWorkbook().save("UsingGenericList_out.xlsx");
		// ExEnd: 1
	}
}