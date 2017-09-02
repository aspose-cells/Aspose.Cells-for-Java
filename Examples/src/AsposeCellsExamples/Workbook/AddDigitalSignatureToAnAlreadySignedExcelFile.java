package AsposeCellsExamples.Workbook;

import java.io.*;
import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AddDigitalSignatureToAnAlreadySignedExcelFile {

	public static void main(String[] args) throws Exception {


		HtmlSaveOptions o=new HtmlSaveOptions();
		o.setExportComments(true);
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();
		
		// Certificate file and its password
		String certFileName = "AsposeTest.pfx";
		String password = "aspose";
		 
		// Load the workbook which is already digitally signed to add new digital signature
		Workbook workbook = new Workbook(srcDir + "sampleDigitallySignedByCells.xlsx");
		 
		// Create the digital signature collection
		DigitalSignatureCollection dsCollection = new DigitalSignatureCollection();
		 
		// Create new digital signature and add it in digital signature collection
		// ------------------------------------------------------------
		// --------------Begin::creating signature---------------------
		 
		// Load the certificate into an instance of InputStream
		InputStream inStream = new FileInputStream(srcDir + certFileName);
		 
		// Create an instance of KeyStore with PKCS12 cryptography
		java.security.KeyStore inputKeyStore = java.security.KeyStore.getInstance("PKCS12");
		 
		// Use the KeyStore.load method to load the certificate stream and its password
		inputKeyStore.load(inStream, password.toCharArray());
		 
		// Create an instance of DigitalSignature and pass the instance of KeyStore, password, comments and time
		DigitalSignature signature = new DigitalSignature(inputKeyStore, password,
		        "Aspose.Cells added new digital signature in existing digitally signed workbook.",
		        com.aspose.cells.DateTime.getNow());
		 
		dsCollection.add(signature);
		// ------------------------------------------------------------
		// --------------End::creating signature-----------------------
		 
		// Add digital signature collection inside the workbook
		workbook.addDigitalSignature(dsCollection);
		 
		// Save the workbook and dispose it.
		workbook.save(outDir + "outputDigitallySignedByCells.xlsx");
		workbook.dispose();

		// Print the message
		System.out.println("AddDigitalSignatureToAnAlreadySignedExcelFile executed successfully.");
	}
}
