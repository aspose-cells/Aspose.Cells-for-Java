package com.aspose.gridweb.test.data;
 
import java.io.Serializable;

import com.aspose.gridweb.CellEventArgs;
import com.aspose.gridweb.CellEventHandler;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
public   class SortEventHandler  implements CellEventHandler,Serializable{

	/**
	 * 
	 */
	private static final long serialVersionUID = -2054739168408657526L;
    private boolean isacend=false;
	@Override
	public void handleCellEvent(Object sender, CellEventArgs e) {
		System.out.println(" handleCellEvent  ...."+isacend);
		GridWebBean gridweb=(GridWebBean)sender;
		  GridWorksheet sheet = gridweb.getWorkSheets().get(0);
		  GridWorksheet sheet1 = gridweb.getWorkSheets().get(1);
		if (e.getArgument().toString().equals("A1")) {
			//reverse order while click again
			isacend=!isacend;
			sheet.getCells().sort(1, 0, 20, 4, 0, isacend,true,false);
		} else if (e.getArgument().toString().equals("B1")) {
			sheet.getCells().sort(1, 0, 20, 4, 1, true,true,false);
		} else if (e.getArgument().toString().equals("C1")) {
			sheet.getCells().sort(1, 0, 20, 4, 2, true,true,false);
		} else if (e.getArgument().toString().equals("D1")) {
			sheet.getCells().sort(1, 0, 20, 4, 3, true,true,false);
		} else if (e.getArgument().toString().equals("1A1")) {
			
			sheet1.getCells().sort(0, 1, 4, 7, 0, true,true,true);
		} else if (e.getArgument().toString().equals("1A2")) {
			sheet1.getCells().sort(0, 1, 4, 7, 1, true,true,true);
		} else if (e.getArgument().toString().equals("1A3")) {
			sheet1.getCells().sort(0, 1, 4, 7, 2, true,true,true);
		} else if (e.getArgument().toString().equals("1A4")) {
			sheet1.getCells().sort(0, 1, 4, 7, 3, true,true,true);
		}
	}

}
