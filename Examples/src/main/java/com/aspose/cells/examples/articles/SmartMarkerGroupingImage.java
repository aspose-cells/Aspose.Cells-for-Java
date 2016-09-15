package com.aspose.cells.examples.articles;

import java.nio.file.*;
import java.util.ArrayList;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SmartMarkerGroupingImage {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SmartMarkerGroupingImage.class) + "articles/";

		SmartMarkerGroupingImage grouping = new SmartMarkerGroupingImage();
		grouping.Execute(dataDir);
	}

	public void Execute(String dataDir) throws Exception {

		// Get the image
		Path path = Paths.get(dataDir + "sample1.png");
		byte[] photo1 = Files.readAllBytes(path);

		// Get the image
		path = Paths.get(dataDir + "sample2.jpg");
		byte[] photo2 = Files.readAllBytes(path);

		// Create a new workbook and access its worksheet
		Workbook workbook = new Workbook();
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Set the standard row height to 35
		worksheet.getCells().setStandardHeight(35);

		// Set column widhts of D, E and F
		worksheet.getCells().setColumnWidth(3, 20);
		worksheet.getCells().setColumnWidth(4, 20);
		worksheet.getCells().setColumnWidth(5, 40);

		// Add the headings in columns D, E and F
		worksheet.getCells().get("D1").putValue("Name");
		Style st = worksheet.getCells().get("D1").getStyle();
		st.getFont().setBold(true);
		worksheet.getCells().get("D1").setStyle(st);

		worksheet.getCells().get("E1").putValue("City");
		worksheet.getCells().get("E1").setStyle(st);

		worksheet.getCells().get("F1").putValue("Photo");
		worksheet.getCells().get("F1").setStyle(st);

		// Add smart marker tags in columns D, E, F
		worksheet.getCells().get("D2").putValue("&=Person.Name(group:normal,skip:1)");
		worksheet.getCells().get("E2").putValue("&=Person.City");
		worksheet.getCells().get("F2").putValue("&=Person.Photo(Picture:FitToCell)");

		// Create Persons objects with photos
		ArrayList<Person> persons = new ArrayList<Person>();
		persons.add(new Person("George", "New York", photo1));
		persons.add(new Person("George", "New York", photo2));
		persons.add(new Person("George", "New York", photo1));
		persons.add(new Person("George", "New York", photo2));
		persons.add(new Person("Johnson", "London", photo2));
		persons.add(new Person("Johnson", "London", photo1));
		persons.add(new Person("Johnson", "London", photo2));
		persons.add(new Person("Simon", "Paris", photo1));
		persons.add(new Person("Simon", "Paris", photo2));
		persons.add(new Person("Simon", "Paris", photo1));
		persons.add(new Person("Henry", "Sydney", photo2));
		persons.add(new Person("Henry", "Sydney", photo1));
		persons.add(new Person("Henry", "Sydney", photo2));

		// Create a workbook designer
		WorkbookDesigner designer = new WorkbookDesigner(workbook);

		// Set the data source and process smart marker tags
		designer.setDataSource("Person", persons);
		designer.process();

		// Save the workbook
		workbook.save(dataDir + "output.xlsx", SaveFormat.XLSX);

		System.out.println("File saved");
	}

	public class Person {
		// Create Name, City and Photo properties
		private String m_Name;
		private String m_City;
		private byte[] m_Photo;

		public Person(String name, String city, byte[] photo) {
			m_Name = name;
			m_City = city;
			m_Photo = photo;
		}

		public String getName() {
			return this.m_Name;
		}

		public void setName(String name) {
			this.m_Name = name;
		}

		public String getCity() {
			return this.m_City;
		}

		public void setCity(String address) {
			this.m_City = address;
		}

		public byte[] getPhoto() {
			return this.m_Photo;
		}

		public void setAddress(byte[] photo) {
			this.m_Photo = photo;
		}
	}

}