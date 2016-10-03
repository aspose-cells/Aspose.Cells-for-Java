package com.aspose.cells.examples.SmartMarkers;

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
