package com.aspose.gridjs.demo;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.stereotype.Component;
import org.springframework.context.ApplicationContext;

import com.aspose.cells.Workbook;
import com.aspose.gridjs.Config;
import com.aspose.gridjs.GridJsWorkbook;

import java.io.File;

import javax.annotation.PostConstruct;


@Component
class MyConfig {

	@Value("${testconfig.CachePath}")
	public String cachePath;

    
	@Value("${testconfig.AsposeLicensePath}")
	public String asposeLicensePath;
	
	
}

@SpringBootApplication
public class GridjsdemoApplication {
	 



	//settings for GridJs，implement GridCacheForStream to store cache file 
	private  static void init(MyConfig myConfig) {
		
		 
	 
		try {
			Config.setFileCacheDirectory(myConfig.cachePath);
			//lazy loading
			Config.setLazyLoading(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
//      when use GridJsWorkbook.CacheImp,no need to setImageUrlBase 		
//		GridJsWorkbook.setImageUrlBase("/GridJs2/Image");
		
	    LocalFileCache mwc = new LocalFileCache();
        GridJsWorkbook.CacheImp = mwc;
        GridJsWorkbook.UpdateMonitor = new ModifyMonitor();
	}
	
	//settings for GridJs,just set a temp path to store cache file 
	//simple way not use GridJsWorkbook.CacheImp ,it shall also ok
	private static void init2(MyConfig myConfig) {
		 
		try {
			Config.setFileCacheDirectory(myConfig.cachePath);
			//lazy loading
			Config.setLazyLoading(true);
 
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		//need to setImageUrlBase in this case
 		GridJsWorkbook.setImageUrlBase("/GridJs2/Image");
 		GridJsWorkbook.UpdateMonitor = new ModifyMonitor();

	}

	
	public static void main(String[] args) {
		
		
 
		 ApplicationContext context  = 	SpringApplication.run(GridjsdemoApplication.class, args);
		 
		 MyConfig myConfig = context.getBean(MyConfig.class);
		
		 //set license for Aspose.Cells
		 com.aspose.cells.License  lic=new com.aspose.cells.License();
			if ((new File(myConfig.asposeLicensePath)).exists()) {
				lic.setLicense(myConfig.asposeLicensePath);
			}
		 init(myConfig);
 
	}

}
