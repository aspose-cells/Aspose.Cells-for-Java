package AsposeCellsExamples.CellsHelperClass;
import com.aspose.cells.CellsHelper;

public class CreateSafeSheetNames {


	public static void main(String[] args) throws Exception {

		// ExStart:CreateSafeSheetNames
        // Long name will be truncated to 31 characters
        String name1 = CellsHelper.createSafeSheetName("this is first name which is created using CellsHelper.CreateSafeSheetName and truncated to 31 characters");

        // Any invalid character will be replaced with _
        String name2 = CellsHelper.createSafeSheetName(" <> + (adj.Private ? \" Private\" : \")", '_');//? shall be replaced with _

        // Display first name
        System.out.println(name1);

        // Display second name
        System.out.println(name2);
        
		// Print message
		System.out.println("Create Safe Sheet Names performed successfully.");
        
        // ExEnd:CreateSafeSheetNames		
	}
}
