package com.aspose.gridweb.test.data;

import java.io.Serializable;

import com.aspose.gridweb.CellEventArgs;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.WorkbookEventHandler;
 

public class MsgPutWorkbookHandler implements WorkbookEventHandler,Serializable{
	public static void setMessageInCell(Object sender,String msg) {
		GridWebBean gridweb = (GridWebBean) sender;
		int row = (gridweb.getCurrentPageIndex()) * gridweb.getPageSize();
		GridWorksheet sheet = gridweb.getActiveSheet();
		sheet.getCells().get(row, 0).setValue(msg);
		sheet.getCells().setColumnWidthPixel(0, 180);
	}
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	private String msg;
	public MsgPutWorkbookHandler(String msg) {
		super();
		this.msg = msg;
	}
	@Override
	public void handleCellEvent(Object sender, CellEventArgs e) {
		setMessageInCell(sender,msg);
		
	}

}
