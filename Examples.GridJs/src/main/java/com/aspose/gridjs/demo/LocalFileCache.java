package com.aspose.gridjs.demo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.gridjs.Config;
import com.aspose.gridjs.GridCacheForStream;

public class LocalFileCache extends GridCacheForStream {
	 
	private String streamcacheDir=null;
	//default constructor
	public LocalFileCache() {
		//make sure the cache path existed
		streamcacheDir=Paths.get(Config.getFileCacheDirectory(), "streamcache").toString();
		File dir = new File(streamcacheDir);
        if (!dir.exists()) {
            try {
				Files.createDirectories(Paths.get(streamcacheDir));
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        }
	}
	
	@Override
	public void saveStream(InputStream s, String uid) {
		// save stream to the cache path with uid
		String filepath = Paths.get(streamcacheDir, uid.replace('/', '.')).toString();
		try (FileOutputStream fos = new FileOutputStream(filepath.toString())) {
			s.reset(); // Equivalent to s.Position = 0 in C#
			byte[] buffer = new byte[1024];
			int length;
			while ((length = s.read(buffer)) > 0) {
				fos.write(buffer, 0, length);
			}
			fos.flush();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	@Override
	public InputStream loadStream(String uid) {
		//load file from the cache path by uid
		String filepath = Paths.get(streamcacheDir, uid.replace('/', '.')).toString();
		try {
			//here you can update the stream per your business logic if you wish
			//for example below we modify the cell a1 value to hello  
			/*
			Workbook wb = new Workbook(filepath);
			wb.getWorksheets().get(0).getCells().get("A1").putValue("hello");
			ByteArrayOutputStream bos = new ByteArrayOutputStream();
			wb.save(bos, SaveFormat.XLSX);
			byte[] byteArray = bos.toByteArray();
			bos.close();  
			ByteArrayInputStream bis = new ByteArrayInputStream(byteArray);
			return bis;
			*/
			return new FileInputStream(filepath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return null;
		}
	}

	@Override
	public boolean isExisted(String uid) {
		//check if the file with uid existed in the cache path
		String filepath = Paths.get(streamcacheDir, uid.replace('/', '.')).toString();
		return Files.exists(Paths.get(filepath));
	}

	//return the action Url to get the file from the cache path by uid
	@Override
	public String getFileUrl(String uid) {
		return "/GridJs2/GetFileUseCacheStream?id=" + uid;
	}

}
