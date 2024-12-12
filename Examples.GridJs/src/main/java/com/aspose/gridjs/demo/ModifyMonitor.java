package com.aspose.gridjs.demo;

import java.util.ArrayList;

import com.aspose.cells.Workbook;

public class ModifyMonitor extends com.aspose.gridjs.GridUpdateMonitor{

	@Override
	public void afterUpdate(String operation, String uid, ArrayList cells) {
		System.out.println("afterUpdate operation is:"+operation+" ,uid is:"+uid+",modified cells count:"+cells.size());
		
	}

	@Override
	public void beforeUpdate(String operation, String uid, Workbook wb) {
		System.out.println("beforeUpdate operation is:"+operation+" ,uid is:"+uid+",workbook sheets count:"+wb.getWorksheets().getActiveSheetName());
		
	}

}
