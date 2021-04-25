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


	public void inserColumn(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int columnIndex = Integer.parseInt(request.getParameter("columnIndex"));
		GridCells gridCells = gridweb.getActiveSheet().getCells();
		gridCells.insertColumn(columnIndex);
	}

	public void deleteColumn(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int columnIndex = Integer.parseInt(request.getParameter("columnIndex"));
		GridCells gridCells = gridweb.getActiveSheet().getCells();
		gridCells.deleteColumn(columnIndex);
	}

	public void insertRow(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int rowIndex = Integer.parseInt(request.getParameter("rowIndex"));
		GridCells gridCells = gridweb.getActiveSheet().getCells();
		gridCells.insertRow(rowIndex);
		//getGridCells(gridweb,request).insertRow(rowIndex);
	}

	public void deleteRow(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int rowIndex = Integer.parseInt(request.getParameter("rowIndex"));
		GridCells gridCells = gridweb.getActiveSheet().getCells();
		gridCells.deleteRow(rowIndex);
	}

	public void mergeCells(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int startRow = Integer.parseInt(request.getParameter("startRow"));
		int startColumn = Integer.parseInt(request.getParameter("startColumn"));
		int rowNumber = Integer.parseInt(request.getParameter("rowNumber"));
		int columnNumber = Integer.parseInt(request.getParameter("columnNumber"));
		GridCells gridCells = gridweb.getActiveSheet().getCells();
		gridCells.merge(startRow, startColumn, rowNumber, columnNumber);
	}

	public void addComment(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		int startRow_c = Integer.parseInt(request.getParameter("startRow_c"));
		int startColumn_c = Integer.parseInt(request.getParameter("startColumn_c"));
		String comment = request.getParameter("comment");
		GridWorksheet gridWorksheet = gridweb.getActiveSheet();
		GridCommentCollection gridCommentCollection = gridWorksheet.getComments();
		gridCommentCollection.add(startRow_c, startColumn_c);
		GridComment gridComment = gridCommentCollection.get(startRow_c, startColumn_c);
		gridComment.setNote(comment);
	}

	public void removeComment(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		int startRow_c = Integer.parseInt(request.getParameter("startRow_c"));
		int startColumn_c = Integer.parseInt(request.getParameter("startColumn_c"));
		GridWorksheet gridWorksheet = gridweb.getActiveSheet();
		gridWorksheet.getComments().removeAt(startRow_c, startColumn_c);
	}
}
