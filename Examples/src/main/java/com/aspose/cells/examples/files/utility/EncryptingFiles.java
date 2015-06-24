/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.files.utility;

import com.aspose.cells.EncryptionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class EncryptingFiles {

    public static void main(String[] args) throws Exception {
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
        workbook.save(dataDir + "encryptedBook1.xls");

        // Print message
        System.out.println("Encryption applied successfully on output file.");
    }
}
