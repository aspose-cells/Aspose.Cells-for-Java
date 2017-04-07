package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class CustomizeGlobalizationSettingsforPivotTable {
	
	class CustomPivotTableGlobalizationSettings extends GlobalizationSettings
	{   
	    //Gets the name of "Total" label in the PivotTable.
	    //You need to override this method when the PivotTable contains two or more PivotFields in the data area.
	    public String getPivotTotalName()
	    {
	        System.out.println("---------GetPivotTotalName-------------");
	        return "AsposeGetPivotTotalName";
	    }
	  
	    //Gets the name of "Grand Total" label in the PivotTable.
	    public String getPivotGrandTotalName()
	    {
	        System.out.println("---------GetPivotGrandTotalName-------------");
	        return "AsposeGetPivotGrandTotalName";
	    }
	  
	    //Gets the name of "(Multiple Items)" label in the PivotTable.
	    public String getMultipleItemsName()
	    {
	        System.out.println("---------GetMultipleItemsName-------------");
	        return "AsposeGetMultipleItemsName";
	    }
	  
	    //Gets the name of "(All)" label in the PivotTable.
	    public String getAllName()
	    {
	        System.out.println("---------GetAllName-------------");
	        return "AsposeGetAllName";
	    }
	  
	    //Gets the name of "Column Labels" label in the PivotTable.
	    public String getColumnLablesName()
	    {
	        System.out.println("---------GetColumnLablesName-------------");
	        return "AsposeGetColumnLablesName";
	    }
	  
	    //Gets the name of "Row Labels" label in the PivotTable.
	    public String getRowLablesName()
	    {
	        System.out.println("---------GetRowLablesName-------------");
	        return "AsposeGetRowLablesName";
	    }
	  
	    //Gets the name of "(blank)" label in the PivotTable.
	    public String getEmptyDataName()
	    {
	        System.out.println("---------GetEmptyDataName-------------");
	        return "(blank)AsposeGetEmptyDataName";
	    }
	  
	    //Gets the name of PivotFieldSubtotalType type in the PivotTable.
	    public String getSubTotalName(int subTotalType)
	    {
	        System.out.println("---------GetSubTotalName-------------");
	  
	        switch (subTotalType)
	        {
	            case PivotFieldSubtotalType.SUM:
	                return "AsposeSum";//polish
	  
	            case PivotFieldSubtotalType.COUNT:
	                return "AsposeCount";
	  
	            case PivotFieldSubtotalType.AVERAGE:
	                return "AsposeAverage";
	  
	            case PivotFieldSubtotalType.MAX:
	                return "AsposeMax";
	  
	            case PivotFieldSubtotalType.MIN:
	                return "AsposeMin";
	  
	            case PivotFieldSubtotalType.PRODUCT:
	                return "AsposeProduct";
	  
	            case PivotFieldSubtotalType.COUNT_NUMS:
	                return "AsposeCount";
	  
	            case PivotFieldSubtotalType.STDEV:
	                return "AsposeStdDev";
	  
	            case PivotFieldSubtotalType.STDEVP:
	                return "AsposeStdDevp";
	  
	            case PivotFieldSubtotalType.VAR:
	                return "AsposeVar";
	  
	            case PivotFieldSubtotalType.VARP:
	                return "AsposeVarp";
	  
	        }
	  
	        return "AsposeSubTotalName";
	    }
	 
	}//End CustomPivotTableGlobalizationSettings

	public void RunCustomizeGlobalizationSettingsforPivotTable() throws Exception
	{
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CustomizeGlobalizationSettingsforPivotTable.class) + "PivotTables/";
		
		//Load your excel file
		Workbook wb = new Workbook(dataDir + "samplePivotTableGlobalizationSettings.xlsx");
		  
		//Setting Custom Pivot Table Globalization Settings
		wb.getSettings().setGlobalizationSettings(new CustomPivotTableGlobalizationSettings());
		  
		//Hide first worksheet that contains the data of the pivot table
		wb.getWorksheets().get(0).setVisible(false);
		  
		//Access second worksheet
		Worksheet ws = wb.getWorksheets().get(1);
		  
		//Access the pivot table, refresh and calculate its data
		PivotTable pt = ws.getPivotTables().get(0);
		pt.setRefreshDataFlag(true);
		pt.refreshData();
		pt.calculateData();
		pt.setRefreshDataFlag(false);
		  
		//Pdf save options - save entire worksheet on a single pdf page
		PdfSaveOptions options = new PdfSaveOptions();
		options.setOnePagePerSheet(true);
		  
		//Save the output pdf 
		wb.save(dataDir + "outputPivotTableGlobalizationSettings.pdf", options);

	}
	
	
	
	public static void main(String[] args) throws Exception {
		CustomizeGlobalizationSettingsforPivotTable pg = new CustomizeGlobalizationSettingsforPivotTable();
		pg.RunCustomizeGlobalizationSettingsforPivotTable();
	}
}
