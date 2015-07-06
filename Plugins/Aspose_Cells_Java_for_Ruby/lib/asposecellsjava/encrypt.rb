module Asposecellsjava
  module Encrypt
    def initialize()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Instantiating a Workbook object by excel file path
        workbook = Rjb::import('com.aspose.cells.Workbook').new(data_dir + 'Book1.xls')
        
        # Password protect the file.
        workbook.getSettings().setPassword("1234")

        encryption_type = Rjb::import('com.aspose.cells.EncryptionType')        

        # Specify XOR encrption type.
        workbook.setEncryptionOptions(encryption_type.XOR, 40)

        # Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
        workbook.setEncryptionOptions(encryption_type.STRONG_CRYPTOGRAPHIC_PROVIDER, 128)
        
        # Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(data_dir + "encrypt.xls")

        puts "Apply encryption, please check the output file."
    end
  end
end
