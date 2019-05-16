package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AddWebExtension {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the output directory.
		String outDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook();

		WebExtensionCollection extensions = workbook.getWorksheets().getWebExtensions();
		WebExtensionTaskPaneCollection taskPanes = workbook.getWorksheets().getWebExtensionTaskPanes();
		
		int extensionIndex = extensions.add();
        int taskPaneIndex = taskPanes.add();
		
        WebExtension extension = extensions.get(extensionIndex);
        extension.getReference().setId("wa104379955");
        extension.getReference().setStoreName("en-US");
        extension.getReference().setStoreType(WebExtensionStoreType.OMEX);
        
        WebExtensionTaskPane taskPane = taskPanes.get(taskPaneIndex);
        taskPane.setVisible(true);
        taskPane.setDockState("right");
        taskPane.setWebExtension(extension);
        
        workbook.save(outDir + "AddWebExtension_Out.xlsx");
        // ExEnd:1

		System.out.println("AddWebExtension executed successfully.");
	}
}
