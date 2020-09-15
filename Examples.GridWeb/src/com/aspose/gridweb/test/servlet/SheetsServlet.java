package com.aspose.gridweb.test.servlet;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.gridweb.Color;
import com.aspose.gridweb.FontUnit;
import com.aspose.gridweb.GridCell;
import com.aspose.gridweb.GridCells;
import com.aspose.gridweb.GridTableItemStyle;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.GridWorksheetCollection;
import com.aspose.gridweb.HorizontalAlign;
import com.aspose.gridweb.Unit;
import com.aspose.gridweb.test.TestGridWebBaseServlet;

/**
 * import sheets.jsp
 */
public class SheetsServlet extends TestGridWebBaseServlet {
	private static final long serialVersionUID = 1L;

	// Add
	public void add(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		// GridWorksheetCollection gridWorksheetCollection = gridweb
		// .getWorkSheets();
		// int index = gridWorksheetCollection.getCount() + 1;
		// gridWorksheetCollection.add("Sheet" + index);
		// gridweb.setActiveSheetIndex(index);

		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		int index= gridWorksheetCollection.add();
		setNameByCount(gridWorksheetCollection, index,"sheet");
		gridweb.setActiveSheetIndex(index);

	}

	// Add Copy
	public void copy(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) throws Exception {
		
		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		int index = gridWorksheetCollection.addCopy(gridweb.getActiveSheetIndex());
		setNameByCount(gridWorksheetCollection, index,"copysheet");
		gridweb.setActiveSheetIndex(index);
	}

	private void setNameByCount(GridWorksheetCollection gridWorksheetCollection, int index,String base) {
		GridWorksheet gw = gridWorksheetCollection.get(index);
		int i = gridWorksheetCollection.getCount();
		gw.setName(base+i);
	}

	// Remove Active Sheet
	public void remove(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) throws Exception {
		
		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		gridWorksheetCollection.removeAt(gridweb.getActiveSheetIndex());
	}

	// Reload data
	@Override
	public void reload(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
	 
		InitData(gridweb,request);
		
		gridweb.setActiveSheetIndex(0);

		
		}

	
	 private void InitData(GridWebBean gridweb,HttpServletRequest request)
	  { 
	  
		 GridWorksheetCollection sheets = gridweb.getWorkSheets();
		 sheets.clear();
		// gridweb..Clear();
	    GridWorksheet sheet =sheets.add("Students");
	    GridCells cells = sheet.getCells();
	    GridCell cell00=cells.getCell(0, 0);
	    cell00.putValue("Name");
	    GridTableItemStyle style=cell00.getStyle();
	    style.getFont().setSize(FontUnit.Point(10));//.Font.Size = new FontUnit("10pt");
	    style.getFont().setBold(true);
	    style.setForeColor(Color.getBlack());
	    style.setHorizontalAlign(HorizontalAlign.Center);
	    style.setBorderWidth(Unit.Pixel(1));
	    cell00.setStyle(style);
	    
	    GridCell cell01=cells.getCell(0, 1);
	    cell01.putValue("Gender");
	    cell01.setStyle(style);

	    GridCell cell02=cells.getCell(0, 2);
	    cell02.putValue("Age");
	    cell02.setStyle(style);

	    GridCell cell03=cells.getCell(0, 3);
	    cell03.putValue("Class");
	    cell03.setStyle(style);

	    cells.getCell(1, 0).putValue("Jack");
	    cells.getCell(1, 1).putValue("M");
	    cells.getCell(1, 2).putValue(19);
	    cells.getCell(1, 3).putValue("One");

	    cells.getCell(2, 0).putValue("Tome");
	    cells.getCell(2, 1).putValue("M");
	    cells.getCell(2, 2).putValue(20);
	    cells.getCell(2, 3).putValue("Four");

	    cells.getCell(3, 0).putValue("Jeney");
	    cells.getCell(3, 1).putValue("W");
	    cells.getCell(3, 2).putValue(18);
	    cells.getCell(3, 3).putValue("Two");

	    cells.getCell(4, 0).putValue("Marry");
	    cells.getCell(4, 1).putValue("W");
	    cells.getCell(4, 2).putValue(17);
	    cells.getCell(4, 3).putValue("There");

	    cells.getCell(5, 0).putValue("Amy");
	    cells.getCell(5, 1).putValue("W");
	    cells.getCell(5, 2).putValue(16);
	    cells.getCell(5, 3).putValue("Four");

	    cells.getCell(6, 0).putValue("Ben");
	    cells.getCell(6, 1).putValue("M");
	    cells.getCell(6, 2).putValue(17);
	    cells.getCell(6, 3).putValue("Four");

	    cells.setColumnWidth(0, 10);
	    cells.setColumnWidth(1, 10);
	    cells.setColumnWidth(2, 10);
	    cells.setColumnWidth(3, 10);
	}

}
