package com.aspose.cells.examples.articles;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.LightCellsDataProvider;
import com.aspose.cells.Row;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;

public class LightCellsDataProviderDemo implements LightCellsDataProvider {

	private final int sheetCount;
	private final int maxRowIndex;
	private final int maxColIndex;
	private int rowIndex;
	private int colIndex;
	private final Style style1;
	private final Style style2;

	public LightCellsDataProviderDemo(Workbook wb, int sheetCount, int rowCount, int colCount) {
		// set the variables/objects
		this.sheetCount = sheetCount;
		this.maxRowIndex = rowCount - 1;
		this.maxColIndex = colCount - 1;

		// add new style object with specific formattings
		int index = wb.getStyles().add();
		style1 = wb.getStyles().get(index);
		Font font = style1.getFont();
		font.setName("MS Sans Serif");
		font.setSize(10);
		font.setBold(true);
		font.setItalic(true);
		font.setUnderline(FontUnderlineType.SINGLE);
		font.setColor(Color.fromArgb(0xffff0000));
		style1.setHorizontalAlignment(TextAlignmentType.CENTER);

		// create another style
		index = wb.getStyles().add();
		style2 = wb.getStyles().get(index);
		style2.setCustom("#,##0.00");
		font = style2.getFont();
		font.setName("Copperplate Gothic Bold");
		font.setSize(8);
		style2.setPattern(BackgroundType.SOLID);
		style2.setForegroundColor(Color.fromArgb(0xff0000ff));
		style2.setBorder(BorderType.TOP_BORDER, CellBorderType.THICK, Color.getBlack());
		style2.setVerticalAlignment(TextAlignmentType.CENTER);
	}

	public boolean isGatherString() {
		return false;
	}

	public int nextCell() {
		if (colIndex < maxColIndex) {
			colIndex++;
			return colIndex;
		}
		return -1;
	}

	public int nextRow() {
		if (rowIndex < maxRowIndex) {
			rowIndex++;
			colIndex = -1; // reset column index
			if (rowIndex % 1000 == 0) {
				System.out.println("Row " + rowIndex);
			}
			return rowIndex;
		}
		return -1;
	}

	public void startCell(Cell cell) {
		if (rowIndex % 50 == 0 && (colIndex == 0 || colIndex == 3)) {
			// do not change the content of hyperlink.
			return;
		}
		if (colIndex < 10) {
			cell.putValue("test_" + rowIndex + "_" + colIndex);
			cell.setStyle(style1);
		} else {
			if (colIndex == 19) {
				cell.setFormula("=Rand() + test!L1");
			} else {
				cell.putValue(rowIndex * colIndex);
			}
			cell.setStyle(style2);
		}
	}

	public void startRow(Row row) {
		row.setHeight(25);
	}

	public boolean startSheet(int sheetIndex) {
		if (sheetIndex < sheetCount) {
			// reset row/column index
			rowIndex = -1;
			colIndex = -1;
			return true;
		}
		return false;
	}

}
