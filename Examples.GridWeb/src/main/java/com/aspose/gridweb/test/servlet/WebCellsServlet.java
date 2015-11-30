package com.aspose.gridweb.test.servlet;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.gridweb.GridCells;
import com.aspose.gridweb.GridComment;
import com.aspose.gridweb.GridCommentCollection;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.GridWorksheetCollection;
import com.aspose.gridweb.test.TestGridWebBaseServlet;

/**
 * import webcells.jsp
 */
public class WebCellsServlet extends TestGridWebBaseServlet {
	private static final long serialVersionUID = 1L;

	@Override
	public void reload(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
	 
		try {
			super.reloadfile(gridweb,request,"data.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private GridCells getGridCells(GridWebBean gridweb,HttpServletRequest request) {
		
		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		GridCells gridCells = gridWorksheetCollection.get(gridweb.getActiveSheetIndex()).getCells();
		return gridCells;
	}

	public void inserColumn(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int columnIndex = Integer.parseInt(request.getParameter("columnIndex"));
		getGridCells(gridweb,request).insertColumn(columnIndex);
	}

	public void deleteColumn(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int columnIndex = Integer.parseInt(request.getParameter("columnIndex"));
		getGridCells(gridweb,request).deleteColumn(columnIndex);
	}

	public void insertRow(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int rowIndex = Integer.parseInt(request.getParameter("rowIndex"));
		getGridCells(gridweb,request).insertRow(rowIndex);
	}

	public void deleteRow(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int rowIndex = Integer.parseInt(request.getParameter("rowIndex"));
		getGridCells(gridweb,request).deleteRow(rowIndex);
	}

	public void mergeCells(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int startRow = Integer.parseInt(request.getParameter("startRow"));
		int startColumn = Integer.parseInt(request.getParameter("startColumn"));
		int rowNumber = Integer.parseInt(request.getParameter("rowNumber"));
		int columnNumber = Integer.parseInt(request.getParameter("columnNumber"));
		getGridCells(gridweb,request).merge(startRow, startColumn, rowNumber, columnNumber);
	}

	public void addComment(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		int startRow_c = Integer.parseInt(request.getParameter("startRow_c"));
		int startColumn_c = Integer.parseInt(request.getParameter("startColumn_c"));
		String comment = request.getParameter("comment");
		GridWorksheet gridWorksheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());
		GridCommentCollection gridCommentCollection = gridWorksheet.getComments();
		gridCommentCollection.add(startRow_c, startColumn_c);
		GridComment gridComment = gridCommentCollection.get(startRow_c, startColumn_c);
		gridComment.setNote(comment);
	}

	public void removeComment(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		int startRow_c = Integer.parseInt(request.getParameter("startRow_c"));
		int startColumn_c = Integer.parseInt(request.getParameter("startColumn_c"));
		GridWorksheet gridWorksheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());
		gridWorksheet.getComments().removeAt(startRow_c, startColumn_c);
	}
}
