package com.aspose.gridweb.test.data;

import java.io.Serializable;

import com.aspose.gridweb.CustomCommandEventHandler;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
 

public class CUSCommandEventHandler implements CustomCommandEventHandler,Serializable{
	public static void setMessageInCell(Object sender,String msg) {
		GridWebBean gridweb = (GridWebBean) sender;
		int row = (gridweb.getCurrentPageIndex()) * gridweb.getPageSize();
		GridWorksheet sheet = gridweb.getActiveSheet();
		sheet.getCells().get(row, 0).setValue(msg);
	}
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	@Override
	public void handleCellEvent(Object sender, String command) {
		setMessageInCell(sender,command);
		
	}

}
