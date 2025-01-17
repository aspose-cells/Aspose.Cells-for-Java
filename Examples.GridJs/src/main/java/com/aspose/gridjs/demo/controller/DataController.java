package com.aspose.gridjs.demo.controller;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

 
import com.aspose.gridjs.GridJsWorkbook;

@RestController
@RequestMapping({"/gridjsdemo"})
public class DataController {
     @Value("${testconfig.FileName}")
     private String testFileName;//="chart.xlsx";
     
     @Value("${testconfig.ListDir}")
     private String listDir; 
    
    @GetMapping({"/index"})
    public ModelAndView getIndexPage()
    {
    	String uid = GridJsWorkbook.getUidForFile(testFileName);
    	 
    	ModelAndView mv = new ModelAndView();
    	 mv.addObject("uid", uid);
       	 mv.addObject("file", testFileName);
    	 mv.setViewName("index");
    	return mv;
    }
    
    @GetMapping("/list")
	public ModelAndView listFiles() {
		 
		List<String[]> filelist = new ArrayList<>();
		 

		File dir = new File(listDir);
		if (dir.exists() && dir.isDirectory()) {
			File[] files = dir.listFiles();
			if (files != null) {
				for (File file : files) {

					 
                    String filename=file.getName();
					
					//get a unique id for the file
					String uid = GridJsWorkbook.getUidForFile(filename);
					String[] ff={filename,uid};
					filelist.add(ff);
				}
			}
		}

		ModelAndView mv = new ModelAndView();
		 
		mv.addObject("filelist", filelist);
		 
		mv.setViewName("list");
		return mv;
	}
    
    @GetMapping("/Uidtml")
	public ModelAndView uidtml(@RequestParam String filename,@RequestParam String uid) {

		 
    	 
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("uid", uid);
		mv.addObject("file", filename);
    	mv.setViewName("index");
    	return mv;
    }
    
    @Value("${testconfig.UploadPath}")
    private String uploadDir;
    
    @PostMapping("/upload")
    public ModelAndView uploadFile(@RequestParam("file") MultipartFile file) {
        // 检查文件是否为空
        if (file.isEmpty()) {
            return null;
        }

        try {
        	 File dir = new File(uploadDir);
             if (!dir.exists()) {
                 Files.createDirectories(Paths.get(uploadDir));
             }
             
        	 Path filePath = Paths.get(uploadDir, file.getOriginalFilename());
        	 
             file.transferTo(new File(filePath.toString()));

             String filename=file.getOriginalFilename();
				
				//get a unique id for the file
			 String uid = GridJsWorkbook.getUidForFile(filename);
			 
			 ModelAndView mv = new ModelAndView();
		    	mv.addObject("uid", uid);
				mv.addObject("file", filename);
		    	mv.setViewName("fileFromUpload");
		    	return mv;
				
        } catch (Exception e) {
        	e.printStackTrace();
            return null;
        }
    }
    
    
}
 
