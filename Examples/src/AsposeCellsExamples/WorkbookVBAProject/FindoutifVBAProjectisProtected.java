package AsposeCellsExamples.WorkbookVBAProject;

import com.aspose.cells.*;

public class FindoutifVBAProjectisProtected {

	public static void main(String[] args) throws Exception {

		// Create a workbook.
		Workbook wb = new Workbook();

		// Access the VBA project of the workbook.
		VbaProject vbaProj = wb.getVbaProject();

		// Find out if VBA Project is Protected using IsProtected property.
		System.out.println("IsProtected - Before Protecting VBA Project: " + vbaProj.isProtected());

		// Protect the VBA project.
		vbaProj.protect(true, "11");

		// Find out if VBA Project is Protected using IsProtected property.
		System.out.println("IsProtected - After Protecting VBA Project: " + vbaProj.isProtected());

		// Print message
		System.out.println("FindoutifVBAProjectisProtected Done Successfully");

	}
}
