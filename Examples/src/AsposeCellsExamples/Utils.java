package AsposeCellsExamples;

import java.io.File;

public class Utils {

    public static String getDataDir(Class c) {
        File dir = new File(System.getProperty("user.dir"));
        
        System.out.println("shake" + dir.getAbsolutePath());
        
        dir = new File(dir, "src");
        dir = new File(dir, "main");
        dir = new File(dir, "resources");

        for (String s : c.getName().split("\\.")) {
            dir = new File(dir, s);
        }

        if (dir.exists()) {
            System.out.println("Using data directory: " + dir.toString());
        } else {
            dir.mkdirs();
            System.out.println("Creating data directory: " + dir.toString());
        }

        return dir.toString() + File.separator;
    }
    
    public static String getSharedDataDir(Class c) {
        File dir = new File(System.getProperty("user.dir"));
        dir = new File(dir, "src");
        dir = new File(dir, "resources");
        
        return dir.toString() + File.separator;
    }
    
    public static String Get_SourceDirectory()
    {
    	File dir = new File(System.getProperty("user.dir"));
        dir = new File(dir, "src");
        dir = new File(dir, "resources");
        
        String srcDir = dir.toString() + File.separator + "01_SourceDirectory"+ File.separator;

        return srcDir;
    }

    public static String Get_OutputDirectory()
    {
    	File dir = new File(System.getProperty("user.dir"));
        dir = new File(dir, "src");
        dir = new File(dir, "resources");
        
        String outDir = dir.toString()+ File.separator + "02_OutputDirectory"+ File.separator;

        return outDir;
    }
}
