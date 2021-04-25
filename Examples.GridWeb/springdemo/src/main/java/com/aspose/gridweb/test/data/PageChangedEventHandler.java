package com.aspose.gridweb.test.data;

import java.io.Serializable;

import com.aspose.gridweb.CellEventArgs;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.WorkbookEventHandler;
 

public class PageChangedEventHandler  implements WorkbookEventHandler,Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	@Override
	public void handleCellEvent(Object sender, CellEventArgs e) {
		GridWebBean gridweb = (GridWebBean) sender;
		int row=(gridweb.getCurrentPageIndex())*gridweb.getPageSize();
		GridWorksheet sheet = gridweb.getActiveSheet();
		sheet.getCells().get(row,0).setValue("PageIndexChanged"+(gridweb.getCurrentPageIndex()+1));
		
	}

}
