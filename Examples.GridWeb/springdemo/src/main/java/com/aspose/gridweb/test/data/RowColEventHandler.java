package com.aspose.gridweb.test.data;

import java.io.Serializable;

import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.RowColumnEventArgs;
import com.aspose.gridweb.RowColumnEventHandler;
 

public class RowColEventHandler  implements RowColumnEventHandler,Serializable{
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
	private String msg;
	
	public RowColEventHandler(String msg) {
		super();
		this.msg = msg;
	}

	@Override
	public void handleCellEvent(Object sender, RowColumnEventArgs e) {
		String msg2="type:"+e.getType()+",id:"+e.getNum()+ (e.getArgument()!=null?",arg:"+e.getArgument().toString():"");
		setMessageInCell(sender,msg,msg2);
		
	}

}
