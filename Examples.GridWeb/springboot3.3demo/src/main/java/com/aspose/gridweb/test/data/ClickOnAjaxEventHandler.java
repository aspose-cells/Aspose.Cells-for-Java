package com.aspose.gridweb.test.data;

import java.io.Serializable;

import com.aspose.gridweb.CellEventArgs;
import com.aspose.gridweb.CellEventStringHandler;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
 

public class ClickOnAjaxEventHandler implements CellEventStringHandler,Serializable{
	public static void setMessageInCell(Object sender,String msg,String msg2) {
		GridWebBean gridweb = (GridWebBean) sender;
		int row = (gridweb.getCurrentPageIndex()) * gridweb.getPageSize();
		GridWorksheet sheet = gridweb.getActiveSheet();
		sheet.getCells().get(row, 0).setValue(msg);
		sheet.getCells().get(row+1, 0).setValue(msg2);
	}
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	@Override
	public String handleCellEvent(Object sender, CellEventArgs e) {
		setMessageInCell(sender,"CellClickOnAjax",e.toString());
		return e.getCell()+"$$$$_CellEventStringHandler";
	}

}
