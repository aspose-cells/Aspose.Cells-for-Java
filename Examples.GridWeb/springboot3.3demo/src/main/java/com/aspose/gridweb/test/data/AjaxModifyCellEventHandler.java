package com.aspose.gridweb.test.data;

import java.io.Serializable;

import com.aspose.gridweb.CellEventArgs;
import com.aspose.gridweb.CellEventHandler;
import com.aspose.gridweb.GridCell;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;

public class AjaxModifyCellEventHandler  implements CellEventHandler,Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	 
	@Override
	public void handleCellEvent(Object sender, CellEventArgs e) {
		GridWebBean gridweb = (GridWebBean) sender;
		GridWorksheet sheet=gridweb.getActiveSheet();
		GridCell cell=e.getCell();
		  if (cell.getColumn() == 1)
            {
			  GridCell cellToUpdate = sheet.getCells().get(cell.getRow(), cell.getColumn() + 1);

                cellToUpdate.putValue(cell.getValue());
                gridweb.getModifiedCells().add(cellToUpdate);  
            }
		
	}
		
 

}
