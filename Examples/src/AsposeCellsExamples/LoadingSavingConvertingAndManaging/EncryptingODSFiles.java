package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class EncryptingODSFiles {
	
	// The path to the documents directory.
	static String sourceDir = Utils.Get_SourceDirectory();
	static String outputDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		
        //Encrypt an ODS file
        //Encrypted ODS file can only be opened in OpenOffice as Excel does not support encrypted ODS files

        //Initialize loading options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.ODS);

        // Instantiate a Workbook object.
        // Open an ODS file.
        Workbook workbook = new Workbook(sourceDir + "sampleEncryptingODSFiles.ods", loadOptions);

        //Encryption options are not effective for ODS files

        // Password protect the file.
        workbook.getSettings().setPassword("1234");

        // Save the excel file.
        workbook.save(outputDir + "outputEncryptingODSFiles.ods");

        //Decrypt ODS file
        //Decrypted ODS file can be opened both in Excel and OpenOffice          

        // Set original password
        loadOptions.setPassword("1234");

        // Load the encrypted ODS file with the appropriate load options
        Workbook encrypted = new Workbook(outputDir + "outputEncryptingODSFiles.ods", loadOptions);

        // Unprotect the workbook
        encrypted.unprotect("1234");

        // Set the password to null
        encrypted.getSettings().setPassword(null);

        // Save the decrypted ODS file
        encrypted.save(outputDir + "outputDecryptingODSFiles.ods");

		// Print message
		System.out.println("Encryption and Decryption applied successfully on ODS file.");
	}

}
