package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CheckVbaProjectSigned {

	public static void main(String[] args) throws Exception {
		// ExStart:CheckVbaProjectSigned
		String dataDir = Utils.getDataDir(CheckVbaProjectSigned.class);
		String inputPath = dataDir + "Sample1.xlsx";

		Workbook workbook = new Workbook(inputPath);

		System.out.println("VBA Project is Signed: " + workbook.getVbaProject().isSigned());
		// ExEnd:CheckVbaProjectSigned
	}
}
