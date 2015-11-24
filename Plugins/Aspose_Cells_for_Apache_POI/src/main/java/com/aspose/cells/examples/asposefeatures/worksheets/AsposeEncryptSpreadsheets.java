package com.aspose.cells.examples.asposefeatures.worksheets;

import com.aspose.cells.EncryptionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AsposeEncryptSpreadsheets
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeEncryptSpreadsheets.class);
	
	//Instantiate a Workbook object by excel file path
	Workbook workbook = new Workbook(dataDir + "book1.xls");

	//Password protect the file.
	workbook.getSettings().setPassword("1234");

	//Specify XOR encryption type.
	workbook.setEncryptionOptions(EncryptionType.XOR, 40);

	//Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
	workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);

	//Save the excel file.
	workbook.save(dataDir + "AsposeEncryptedWorkBook.xls");
	
	System.out.println("Encryption Done.");
    }
}