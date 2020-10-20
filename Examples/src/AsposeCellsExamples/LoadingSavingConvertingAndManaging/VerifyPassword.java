package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import AsposeCellsExamples.Utils;
import com.aspose.cells.EncryptionType;
import com.aspose.cells.FileFormatUtil;
import com.aspose.cells.Workbook;

import java.io.FileInputStream;

public class VerifyPassword {

	public static void main(String[] args) throws Exception {

		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(VerifyPassword.class) + "LoadingSavingConvertingAndManaging/";

		// Create a Stream object
		FileInputStream fstream = new FileInputStream(dataDir + "EncryptedBook1.xlsx");

		boolean isPasswordValid = FileFormatUtil.verifyPassword(fstream, "1234");

		System.out.println("Password is Valid: " + isPasswordValid);
		// ExEnd:1

		// Print message
		System.out.println("VerifyPassword executed successfully");

	}
}
