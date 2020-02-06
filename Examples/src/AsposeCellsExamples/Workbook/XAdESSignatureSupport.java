package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

import java.io.FileInputStream;
import java.io.InputStream;

public class XAdESSignatureSupport {

	public static void main(String[] args) throws Exception {
        // ExStart: 1
        // directories
        String sourceDir = Utils.Get_SourceDirectory();
        String outputDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "sourceFile.xlsx");
        String password = "pfxPassword";
        String pfx = "pfxFile";

        // Load the certificate into an instance of InputStream
        InputStream inStream = new FileInputStream(pfx);

        // Create an instance of KeyStore with PKCS12 cryptography
        java.security.KeyStore inputKeyStore = java.security.KeyStore.getInstance("PKCS12");

        // Use the KeyStore.load method to load the certificate stream and its password
        inputKeyStore.load(inStream, password.toCharArray());

        DigitalSignature signature = new DigitalSignature(inputKeyStore, password, "testXAdES", com.aspose.cells.DateTime.getNow());

        signature.setXAdESType(XAdESType.X_AD_ES);
        DigitalSignatureCollection dsCollection = new DigitalSignatureCollection();
        dsCollection.add(signature);

        workbook.setDigitalSignature(dsCollection);

        workbook.save(outputDir + "XAdESSignatureSupport_out.xlsx");
        // ExEnd:1

		System.out.println("XAdESSignatureSupport executed successfully.");
	}
}
