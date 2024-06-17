package com.aspose.gridjs.demo.controller;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import com.aspose.gridjs.GridJsWorkbook;

@RestController
@RequestMapping({"/gridjsdemo"})
public class DataController {
     @Value("${testconfig.FileName}")
     private String fileName;//="chart.xlsx";
    
    @GetMapping({"/index"})
    public ModelAndView getIndexPage()
    {
    	String uid = GridJsWorkbook.getUidForFile(fileName);
    	 
    	ModelAndView mv = new ModelAndView();
    	 mv.addObject("uid", uid);
       	 mv.addObject("file", fileName);
    	 mv.setViewName("index");
    	return mv;
    }
    
    
}
 
