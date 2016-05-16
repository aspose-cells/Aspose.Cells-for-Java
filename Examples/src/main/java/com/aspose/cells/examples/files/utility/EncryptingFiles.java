package com.aspose.cells.examples.files.utility;

import com.aspose.cells.EncryptionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class EncryptingFiles {

    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(EncryptingFiles.class);

        //Instantiate a Workbook object by excel file path
        Workbook workbook = new Workbook(dataDir + "Book1.xls");

        //Password protect the file.
        workbook.getSettings().setPassword("1234");

        //Specify XOR encrption type.
        workbook.setEncryptionOptions(EncryptionType.XOR, 40);

        //Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
        workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);

        //Save the excel file.
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Encryption applied successfully on output file.");
        //ExEnd:1
    }
}
