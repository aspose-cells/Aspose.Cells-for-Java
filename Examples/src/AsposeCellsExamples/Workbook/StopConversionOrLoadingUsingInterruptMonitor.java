package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class StopConversionOrLoadingUsingInterruptMonitor 
{
	static String outDir = Utils.Get_OutputDirectory();
	
	//Create InterruptMonitor object
    InterruptMonitor im = new InterruptMonitor();

    public class ThreadStart extends Thread
	{	
		private int ThreadFunc;
		
		public ThreadStart(int threadFunc)
		{
			this.ThreadFunc = threadFunc;
		}
        
        //This function will create workbook and convert it to Pdf format
        void CreateWorkbookAndConvertItToPdfFormat() throws Exception
        {
            //Create a workbook object
            Workbook wb = new Workbook();

            //Assign it InterruptMonitor object
            wb.setInterruptMonitor(im);

            //Access first worksheet
            Worksheet ws = wb.getWorksheets().get(0);

            //Access cell AB1000000 and add some text inside it.
            Cell cell = ws.getCells().get("AB1000000");
            cell.putValue("This is text.");

            try
            {
                //Save the workbook to Pdf format
                wb.save(outDir + "output_InterruptMonitor.pdf");
                
                //Show successfull message
                System.out.println("Excel to Pdf - Successful Conversion");
            }
            catch (CellsException ex)
            {
                System.out.println("Process Interrupted - Message: " + ex.getMessage());
            }
        }
        
        //This function will interrupt the conversion process after 10s
        void WaitForWhileAndThenInterrupt() throws Exception
        {
            Thread.sleep(1000 * 10);
            im.interrupt();
        }
        
        public void run() 
        {
        	try
        	{
        		if(this.ThreadFunc == 1)
        		{
        			CreateWorkbookAndConvertItToPdfFormat();
        		}
            	
            	if(this.ThreadFunc == 2)
            	{
            		WaitForWhileAndThenInterrupt();
            	}
        		
        	}
        	catch(Exception ex)
        	{
        		System.out.println("Process Interrupted - Message: " + ex.getMessage());
        	}
        	
        }
	}//ThreadStart

    public void TestRun() throws Exception
	{
		ThreadStart t1 = new ThreadStart(1);
		ThreadStart t2 = new ThreadStart(2);
		
		t1.start();
		t2.start();
		
		t1.join();
		t2.join();
	}

    public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		new StopConversionOrLoadingUsingInterruptMonitor().TestRun();
		
		// Print the message
		System.out.println("StopConversionOrLoadingUsingInterruptMonitor executed successfully.");
	}

}//StopConversionOrLoadingUsingInterruptMonitor
