package com.aspose.cells.examples.SmartMarkers;

import java.util.ArrayList;

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
