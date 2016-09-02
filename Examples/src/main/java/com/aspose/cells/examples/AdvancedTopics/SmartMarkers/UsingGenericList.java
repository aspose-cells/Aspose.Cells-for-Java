package com.aspose.cells.examples.AdvancedTopics.SmartMarkers;

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

		// Create a designer workbook
		Workbook workbook = new Workbook();

		Worksheet worksheet = workbook.getWorksheets().get(0);

		worksheet.getCells().get("A1").putValue("Husband Name");
		worksheet.getCells().get("A2").putValue("&=Husband.Name");

		worksheet.getCells().get("B1").putValue("Husband Age");
		worksheet.getCells().get("B2").putValue("&=Husband.Age");

		worksheet.getCells().get("C1").putValue("Wife's Name");
		worksheet.getCells().get("C2").putValue("&=Husband.Wives.Name");

		worksheet.getCells().get("D1").putValue("Wife's Age");
		worksheet.getCells().get("D2").putValue("&=Husband.Wives.Age");

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

		ArrayList<Husband> list = new ArrayList<Husband>();

		// Create the relevant Wife objects for the Husband object
		ArrayList<Wife> wives = new ArrayList<Wife>();
		wives.add(new Wife("Chen Zhao", 34));
		wives.add(new Wife("Jamima Winfrey", 28));
		wives.add(new Wife("Reham Smith", 35));

		// Create a Husband object
		Husband h1 = new Husband("Mark John", 30, wives);

		// Create the relevant Wife objects for the Husband object
		wives = new ArrayList<Wife>();
		wives.add(new Wife("Karishma Jathool", 36));
		wives.add(new Wife("Angela Rose", 33));
		wives.add(new Wife("Hina Khanna", 45));

		// Create a Husband object
		Husband h2 = new Husband("Masood Shankar", 40, wives);

		// Add the objects to the list
		list.add(h1);
		list.add(h2);

		// Specify the DataSource
		designer.setDataSource("Husband", list);

		// Process the markers
		designer.process();

		// Autofit columns
		worksheet.autoFitColumns();

		// Save the Excel file.
		designer.getWorkbook().save("output.xlsx");
	}

	// This is the code for Wife.java class
	public class Wife {

		private String m_Name;
		private int m_Age;

		public Wife(String name, int age) {
			this.m_Name = name;
			this.m_Age = age;
		}

		public String getName() {
			return m_Name;
		}

		public int getAge() {
			return m_Age;
		}

	}

	public class Husband {

		private String m_Name;
		private int m_Age;
		private ArrayList<Wife> m_Wives;

		public Husband(String name, int age, ArrayList<Wife> wives) {
			this.m_Name = name;
			this.m_Age = age;
			this.m_Wives = wives;
		}

		public String getName() {
			return m_Name;
		}

		public int getAge() {
			return m_Age;
		}

		public ArrayList<Wife> getWives() {
			return m_Wives;
		}

	}
}
