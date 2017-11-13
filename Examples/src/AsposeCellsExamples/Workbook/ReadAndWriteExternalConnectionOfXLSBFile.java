package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ReadAndWriteExternalConnectionOfXLSBFile { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		//Load the source Excel Xlsb file
		Workbook wb = new Workbook(srcDir + "sampleExternalConnection_XLSB.xlsb");
		  
		//Read the first external connection which is actually a DB-Connection
		DBConnection dbCon = (DBConnection)wb.getDataConnections().get(0);
		  
		//Print the Name, Command and Connection Info of the DB-Connection
		System.out.println("Connection Name: " + dbCon.getName());
		System.out.println("Command: " + dbCon.getCommand());
		System.out.println("Connection Info: " + dbCon.getConnectionInfo());
		  
		//Modify the Connection Name
		dbCon.setName("NewCust");
		  
		//Save the Excel Xlsb file
		wb.save(outDir + "outputExternalConnection_XLSB.xlsb");

		// Print the message
		System.out.println("ReadAndWriteExternalConnectionOfXLSBFile executed successfully.");
	}
}

