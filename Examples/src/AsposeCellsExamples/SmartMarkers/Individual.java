package AsposeCellsExamples.SmartMarkers;

public class Individual extends Person {

	private Person m_Wife;

	public Individual(String name, int age, Person wife) {
	    super(name,age);
		this.m_Wife = wife;
	}

	public Person getWife() {
		return m_Wife;
	}

}
