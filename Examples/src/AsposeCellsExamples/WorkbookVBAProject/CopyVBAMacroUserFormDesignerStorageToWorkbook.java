package AsposeCellsExamples.WorkbookVBAProject;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CopyVBAMacroUserFormDesignerStorageToWorkbook
{

    static String srcDir = Utils.Get_SourceDirectory();
    static String outDir = Utils.Get_OutputDirectory();

    public static void main(String[] args) throws Exception
    {

        System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
        
      //Create empty target workbook
        Workbook target = new Workbook();
         
        //Load the Excel file containing VBA-Macro Designer User Form
        Workbook templateFile = new Workbook(srcDir + "sampleDesignerForm.xlsm");
         
        //Copy all template worksheets to target workboook
        int sheetCount = templateFile.getWorksheets().getCount();

        for(int idx=0; idx<sheetCount; idx++)
        {
        	Worksheet ws = templateFile.getWorksheets().get(idx);
        	
        	if (ws.getType() == SheetType.WORKSHEET)
        	{
        		Worksheet s = target.getWorksheets().add(ws.getName());
        		s.copy(ws);
         
        		//Put message in cell A2 of the target worksheet
        		s.getCells().get("A2").putValue("VBA Macro and User Form copied from template to target.");
        	}
        }//for
        	 
        //-----------------------------------------------
         
        //Copy the VBA-Macro Designer UserForm from Template to Target
        int modCount = templateFile.getWorksheets().getCount();

        for(int idx=0; idx<modCount; idx++)
        {
        	VbaModule vbaItem = templateFile.getVbaProject().getModules().get(idx);
        	
        	if (vbaItem.getName().equals("ThisWorkbook"))
        	{
        		//Copy ThisWorkbook module code
        		target.getVbaProject().getModules().get("ThisWorkbook").setCodes(vbaItem.getCodes());
        	}
        	else
        	{
        		//Copy other modules code and data
        		System.out.println(vbaItem.getName());
         
        		int vbaMod = 0;
        		Worksheet sheet = target.getWorksheets().getSheetByCodeName(vbaItem.getName());
        		if (sheet == null)
        		{
        			vbaMod = target.getVbaProject().getModules().add(vbaItem.getType(), vbaItem.getName());
        		}
        		else
        		{
        			vbaMod = target.getVbaProject().getModules().add(sheet);
        		}
         
        		target.getVbaProject().getModules().get(vbaMod).setCodes(vbaItem.getCodes());
         
        		if ((vbaItem.getType() == VbaModuleType.DESIGNER))
        		{
        			//Get the data of the user form i.e. designer storage
        			byte[] designerStorage = templateFile.getVbaProject().getModules().getDesignerStorage(vbaItem.getName());
         
        			//Add the designer storage to target Vba Project
        			target.getVbaProject().getModules().addDesignerStorage(vbaItem.getName(), designerStorage);
        		}
        	}//else
        }//for
         
        //Save the target workbook
        target.save(outDir + "outputDesignerForm.xlsm", SaveFormat.XLSM);
	
		// Print the message
		System.out.println("CopyVBAMacroUserFormDesignerStorageToWorkbook executed successfully.");
	}
}
