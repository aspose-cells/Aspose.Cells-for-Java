package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CheckVbaProjectSigned {

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(CheckVbaProjectSigned.class) + "articles/";
		String inputPath = dataDir + "Sample1.xlsx";

		Workbook workbook = new Workbook(inputPath);

		System.out.println("VBA Project is Signed: " + workbook.getVbaProject().isSigned());

	}
}
