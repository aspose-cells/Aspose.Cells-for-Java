package com.aspose.gridjsdemo.filemanagement.controller;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
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
     @Value("${fileconfig.FileName}")
     private String testFileName;//="chart.xlsx";
     
     @Value("${fileconfig.ListDir}")
     private String listDir; 
     
     
     @Autowired
     private ResourceLoader resourceLoader;

     @GetMapping("/test-load")
     public ResponseEntity<String> testLoad() {
         try {
             Resource resource = resourceLoader.getResource("classpath:static/test.html");
             if (resource.exists()) {
                 return ResponseEntity.ok("File exists at: " + resource.getURI());
             } else {
                 return ResponseEntity.status(404).body("File not found");
             }
         } catch (Exception e) {
             return ResponseEntity.status(500).body("Error: " + e.getMessage());
         }
     }
     
     
    
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
		 
		mv.setViewName("file/list");
		return mv;
	}
    
    @GetMapping("/Uidtml")
	public ModelAndView uidtml(@RequestParam String filename,@RequestParam String uid) {

		 

		ModelAndView mv = new ModelAndView();
		mv.addObject("uid", uid);
		mv.addObject("file", filename);
		mv.setViewName("file/index");
		return mv;
	}
    
    @Value("${fileconfig.UploadPath}")
    private String uploadDir;
    
    @PostMapping("/upload")
    public ModelAndView uploadFile(@RequestParam("file") MultipartFile file) {
        
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
		    	mv.setViewName("file/fileFromUpload");
		    	return mv;
				
        } catch (Exception e) {
        	e.printStackTrace();
            return null;
        }
    }
    
    
}