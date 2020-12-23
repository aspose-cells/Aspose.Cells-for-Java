package AsposeCellsExamples.SmartMarkers;

import java.util.ArrayList;

public class Teacher extends Person {
	public Teacher(String name, int age, ArrayList<Person> students) {
		super(name, age);
		// TODO Auto-generated constructor stub\
		m_Students = students;
	}

	
	private ArrayList<Person> m_Students;


	public ArrayList<Person> getStudents() {
		return m_Students;
	}
}
