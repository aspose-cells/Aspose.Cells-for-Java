package com.aspose.gridjs.demo;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.aspose.gridjs.Config;
import com.aspose.gridjs.GridJsWorkbook;

import jakarta.annotation.PostConstruct;

@SpringBootApplication
public class GridjsdemoApplication {
	private static String cachePath;

	@Value("${testconfig.CachePath}")
	private String cachePathProperty;

	@PostConstruct
	private void init() {
		cachePath = this.cachePathProperty;
		try {
			Config.setFileCacheDirectory(cachePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
//      when use GridJsWorkbook.CacheImp,no need to setImageUrlBase 		
//		GridJsWorkbook.setImageUrlBase("/GridJs2/Image");
		
	    LocalFileCache mwc = new LocalFileCache();
        GridJsWorkbook.CacheImp = mwc;
	}
	
	//simple way not use GridJsWorkbook.CacheImp ，shall also ok
	private void init2() {
		cachePath = this.cachePathProperty;
		try {
			Config.setFileCacheDirectory(cachePath);
 
		} catch (Exception e) {
			e.printStackTrace();
		}
		//need to setImageUrlBase
 		GridJsWorkbook.setImageUrlBase("/GridJs2/Image");
		
 
	}
	
	public static void main(String[] args) {
		 
		SpringApplication.run(GridjsdemoApplication.class, args);
	}

}
