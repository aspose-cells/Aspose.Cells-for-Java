package com.aspose.cells.examples.articles;

import com.aspose.cells.VbaProject;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddLibraryReference {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddLibraryReference.class) + "articles/";
		Workbook workbook = new Workbook();

		VbaProject vbaProj = workbook.getVbaProject();

		vbaProj.getReferences().addRegisteredReference("stdole",
				"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");
		vbaProj.getReferences().addRegisteredReference("Office",
				"*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE14\\MSO.DLL#Microsoft Office 14.0 Object Library");

		workbook.save(dataDir + "ALReference_out.xlsm");


	}
}
