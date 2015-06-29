package asposefeatures.workingwithworkbook.encryptworkbook.java;

import com.aspose.cells.EncryptionType;
import com.aspose.cells.Workbook;

public class AsposeEncryptSpreadsheets
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/workingwithworkbook/encryptworkbook/data/";
	
	//Instantiate a Workbook object by excel file path
	Workbook workbook = new Workbook(dataPath + "book1.xls");

	//Password protect the file.
	workbook.getSettings().setPassword("1234");

	//Specify XOR encryption type.
	workbook.setEncryptionOptions(EncryptionType.XOR, 40);

	//Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
	workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);

	//Save the excel file.
	workbook.save(dataPath + "AsposeEncryptedWorkBook.xls");
	
	System.out.println("Encryption Done.");
    }
}