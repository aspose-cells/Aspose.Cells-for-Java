package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.CellValueType;
import com.aspose.cells.LightCellsDataHandler;
import com.aspose.cells.Row;
import com.aspose.cells.Worksheet;

public class LightCellsDataHandlerVisitCells implements LightCellsDataHandler {
	public int cellCount;
	public int formulaCount;
	public int stringCount;

	public LightCellsDataHandlerVisitCells() {
		this.cellCount = 0;
		this.formulaCount = 0;
		this.stringCount = 0;
	}

	public int cellCount() {
		return cellCount;
	}

	public int formulaCount() {
		return formulaCount;
	}

	public int stringCount() {
		return stringCount;
	}

	public boolean startSheet(Worksheet sheet) {
		System.out.println("Processing sheet[" + sheet.getName() + "]");
		return true;
	}

	public boolean startRow(int rowIndex) {
		return true;
	}

	public boolean processRow(Row row) {
		return true;
	}

	public boolean startCell(int column) {
		return true;
	}

	public boolean processCell(Cell cell) {
		this.cellCount = this.cellCount + 1;
		if (cell.isFormula()) {
			this.formulaCount = this.formulaCount + 1;
		} else if (cell.getType() == CellValueType.IS_STRING) {
			this.stringCount = this.stringCount + 1;
		}
		return false;
	}
}