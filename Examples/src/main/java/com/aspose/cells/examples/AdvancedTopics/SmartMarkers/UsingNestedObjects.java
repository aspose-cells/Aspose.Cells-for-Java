package com.aspose.cells.examples.AdvancedTopics.SmartMarkers;

import java.util.ArrayList;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;
import com.aspose.cells.examples.Utils;

public class UsingNestedObjects {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(UsingNestedObjects.class);
		Workbook workbook = new Workbook(dataDir + "designer.xlsx");

		WorkbookDesigner designer = new WorkbookDesigner();
		designer.setWorkbook(workbook);

		ArrayList<Individual> list = new ArrayList<Individual>();
		list.add(new Individual("John", 23, new Wife("Jill", 20)));
		list.add(new Individual("Jack", 25, new Wife("Hilly", 21)));
		list.add(new Individual("James", 26, new Wife("Hally", 22)));
		list.add(new Individual("Baptist", 27, new Wife("Newly", 23)));

		designer.setDataSource("Individual", list);

		designer.process(false);

		workbook.save(dataDir + "output.xlsx");
			}

			// This is the code for Individual.java class
			public class Individual {

		private String m_Name;
		private int m_Age;
		private Wife m_Wife;

		public Individual(String name, int age, Wife wife) {
			this.m_Name = name;
			this.m_Age = age;
			this.m_Wife = wife;
		}

		public String getName() {
			return m_Name;
		}

		public int getAge() {
			return m_Age;
		}

		public Wife getWife() {
			return m_Wife;
		}

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
	}
}
