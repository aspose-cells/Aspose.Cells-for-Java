package com.aspose.gridweb.demo.web;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;
 

@RestController
@RequestMapping({"/gridwebdemo"})
public class HelloController {
	 @GetMapping("/hello")  
	    public String hello() {  
	        return "Hello, World!";  
	    }  
    
    @GetMapping({"/sheets"})
    public ModelAndView getSheets()
    {
      ModelAndView mv = new ModelAndView();
      mv.setViewName("sheets");
      return mv;
    }
    @GetMapping({"/createcontent"})
    public ModelAndView getCreateContent()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.setViewName("createcontent");
    	return mv;
    }
 
    @GetMapping({"/chartrefresh"})
    public ModelAndView getChartrefresh()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.setViewName("chartrefresh");
    	return mv;
    }
    @GetMapping({"/modes"})
    public ModelAndView getModes()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.setViewName("modes");
    	return mv;
    }
    @GetMapping({"/conditionformat"})
    public ModelAndView getCondtionformat()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.setViewName("conditionformat");
    	return mv;
    }
    @GetMapping({"/group"})
    public ModelAndView getGroup()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.setViewName("group");
    	return mv;
    }
    @GetMapping({"/controls"})
    public ModelAndView getDisplaycontrols()
    {
      ModelAndView mv = new ModelAndView();
      mv.addObject("servletname", "FeatureServlet");
      mv.addObject("method", "loadControls");
      mv.addObject("title", "displaycontrols");
      mv.addObject("info", "the demo shows the page that loads kinds of controls in  GridWeb");
      mv.setViewName("general");
      return mv;
    }
    
    @GetMapping({"/sort"})
    public ModelAndView getSort()
    {
      ModelAndView mv = new ModelAndView();
      mv.addObject("servletname", "FeatureServlet");
      mv.addObject("method", "sort");
      mv.addObject("title", "sort");
      mv.addObject("info", "The Demo represents the sorting capabilities of GriWeb Control.click the hyperlink on the A1,B1,C1,D1 to do cells sort. ");
      mv.setViewName("general");
      return mv;
    }
    
    @GetMapping({"/autofilter"})
    public ModelAndView getAutofilter()
    {
      ModelAndView mv = new ModelAndView();
      mv.addObject("servletname", "FunctionServlet");
      mv.addObject("method", "autoFilter");
      mv.addObject("title", "auto filter");
      mv.addObject("info", "This Demo Imports an Excel File from a source and Set the AutoFilter feature. ");
      mv.setViewName("general");
      return mv;
    }
    @GetMapping({"/largerows"})
    public ModelAndView getLargerows()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "loadLargeRows");
    	mv.addObject("title", "loadLargeRows");
    	mv.addObject("info", " This demo loads  a file with many rows,every time scroll it will load piece of rows for rendering. ");
    	mv.setViewName("general");
    	return mv;
    }
   
    
    @GetMapping({"/events"})
    public ModelAndView getEvents()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "events");
    	mv.addObject("title", "events");
    	mv.addObject("info", " This Demo Demonstrates Event Handling related to GridWeb Control. ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/clientfunc"})
    public ModelAndView getClientFunc()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "clientfunc");
    	mv.addObject("title", "clientfunc");
    	mv.addObject("info", " trigger client js call back on ajax after cell select,check console out put. ");
    	mv.setViewName("clientfunction");
    	return mv;
    }
    @GetMapping({"/validation"})
    public ModelAndView getValidation()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FunctionServlet");
    	mv.addObject("method", "validation");
    	mv.addObject("title", "validation");
    	mv.addObject("info", " Click <b>Reload</b> to see how demo reloads data and applies validation rules so\r\n" + 
    			"            that invalid (not matching certain RegExp) values could not be entered in the GridWeb\r\n" + 
    			"            Control.");
    	mv.setViewName("validation");
    	return mv;
    }
    @GetMapping({"/webcells"})
    public ModelAndView getWebcells()
    {
    	ModelAndView mv = new ModelAndView();
    	
     
    	mv.addObject("title", "webcells");
    	 
    	mv.setViewName("webcells");
    	return mv;
    }
    @GetMapping({"/commandbutton"})
    public ModelAndView getCommandbutton()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("title", "commandbutton");
    	mv.setViewName("commandbutton");
    	return mv;
    }
    @GetMapping({"/freezepane"})
    public ModelAndView getFreezepane()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "loadFreezePaneFile");
    	mv.addObject("title", "loadFreezePaneFile");
    	mv.addObject("info", "  Click <b>Reload</b> to see how demo loads document, freezes a part of the worksheet\r\n" + 
    			"            so that it remains visible while scrolling and displays data in the GridWeb Control.");
    	mv.setViewName("general");
    	return mv;
    }
   
    @GetMapping({"/freezepane_custom"})
    public ModelAndView getFreezepaneCustom()
    {
    	ModelAndView mv = new ModelAndView();
    	
    	mv.setViewName("freezepane_custom");
    	return mv;
    }
    
    @GetMapping({"/hyperlink"})
    public ModelAndView getHyperlink()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FunctionServlet");
    	mv.addObject("method", "loadHyperlinkFile");
    	mv.addObject("title", "hyperlink");
    	mv.addObject("info", " Click <b>Reload</b> to see how demo demonstrates how to hyperlink table cells so\r\n" + 
    			"            that browser windows would be opened when clicked and displays data in the GridWeb\r\n" + 
    			"            Control.");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/pivot"})
    public ModelAndView getPivot()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "loadPivotFile");
    	mv.addObject("title", "pivot");
    	mv.addObject("info", " this demo shows load pivot file ");
    	mv.setViewName("general");
    	return mv;
    }
    
    @GetMapping({"/customheader"})
    public ModelAndView getCustomheaders()
    {
    	ModelAndView mv = new ModelAndView();
    	 
    	mv.setViewName("customheader");
    	return mv;
    }
    
     @GetMapping({"/math"})
    public ModelAndView getMath()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "loadMathFile");
    	mv.addObject("title", "loadMathFile");
    	mv.addObject("info", "  This demo loads an existing file into an empty WebWorksheet to demonstrate how GridWeb\r\n" + 
    			"            applies typical <b>Math</b> formulas to grid cells and calculates formula values.\r\n" + 
    			"            Click <b>Reload</b> to reload initial data for the grid. You can also Save/Open\r\n" + 
    			"            the output in <b>XLS</b>Format by clicking the Save Button on GridWeb Control Bar.");
    	mv.setViewName("general");
    	return mv;
    }
     
     @GetMapping({"/textdata"})
     public ModelAndView getTextData()
     {
    	 ModelAndView mv = new ModelAndView();
    	 mv.addObject("servletname", "FeatureServlet");
    	 mv.addObject("method", "loadTextAndDataFile");
    	 mv.addObject("title", "loadTextAndDataFile");
    	 mv.addObject("info", "  This demo loads an existing file into an empty WebWorksheet to demonstrate how GridWeb\r\n" + 
    	 		"            applies typical <b>Text</b> and <b>Data</b> formulas to grid cells and calculates\r\n" + 
    	 		"            formula values. Click <b>Reload</b> to reload initial data for the grid. You can\r\n" + 
    	 		"            also Save/Open the output in <b>XLS</b>Format by clicking the Save Button on GridWeb\r\n" + 
    	 		"            Control Bar.");
    	 mv.setViewName("general");
    	 return mv;
     }
     @GetMapping({"/statistical"})
     public ModelAndView getStatistical()
     {
    	 ModelAndView mv = new ModelAndView();
    	 mv.addObject("servletname", "FeatureServlet");
    	 mv.addObject("method", "loadStatisticalFile");
    	 mv.addObject("title", "loadStatisticalFile");
    	 mv.addObject("info", " This demo loads an existing file into an empty WebWorksheet to\r\n" + 
    	 		"			demonstrate how GridWeb applies typical <b>Statistical</b> formulas\r\n" + 
    	 		"			to grid cells and calculates formula values. Click <b>Reload</b> to\r\n" + 
    	 		"			reload initial data for the grid. You can also Save/Open the output\r\n" + 
    	 		"			in <b>XLS</b>Format by clicking the Save Button on GridWeb Control\r\n" + 
    	 		"			Bar.");
    	 mv.setViewName("general");
    	 return mv;
     }
     @GetMapping({"/datetime"})
     public ModelAndView getDatetime()
     {
    	 ModelAndView mv = new ModelAndView();
    	 mv.addObject("servletname", "FeatureServlet");
    	 mv.addObject("method", "loadDateTimeFile");
    	 mv.addObject("title", "loadDateTimeFile");
    	 mv.addObject("info", "This demo loads an existing file into an empty WebWorksheet to demonstrate how GridWeb\r\n" + 
    	 		"            applies typical <b>Date</b> formulas to grid cells and calculates formula values.\r\n" + 
    	 		"            Click <b>Reload</b> to reload initial data for the grid. You can also Save/Open\r\n" + 
    	 		"            the output in <b>XLS</b>Format by clicking the Save Button on GridWeb Control Bar.");
    	 mv.setViewName("general");
    	 return mv;
     }
     @GetMapping({"/logical"})
     public ModelAndView getLogical()
     {
    	 ModelAndView mv = new ModelAndView();
    	 mv.addObject("servletname", "FeatureServlet");
    	 mv.addObject("method", "loadLogicalFile");
    	 mv.addObject("title", "loadLogicalFile");
    	 mv.addObject("info", " This demo loads an existing file into an empty WebWorksheet to demonstrate how GridWeb\r\n" + 
    	 		"            applies typical <b>Logical</b> formulas to grid cells and calculates formula values.\r\n" + 
    	 		"            Click <b>Reload</b> to reload initial data for the grid. You can also Save/Open\r\n" + 
    	 		"            the output in <b>XLS</b>Format by clicking the Save Button on GridWeb Control Bar.");
    	 mv.setViewName("general");
    	 return mv;
     }
     @GetMapping({"/pagination"})
     public ModelAndView getPagination()
     {
    	 ModelAndView mv = new ModelAndView();
    	 mv.addObject("servletname", "FeatureServlet");
    	 mv.addObject("method", "pagination");
    	 mv.addObject("title", "pagination");
    	 mv.addObject("info", " Click <b>Reload</b> to see how demo loads data from data source and\r\n" + 
    	 		"			divides it into several pages to improve performance or support\r\n" + 
    	 		"			logical data division for subsequent data preview in the GridWeb\r\n" + 
    	 		"			Control.");
    	 mv.setViewName("general");
    	 return mv;
     }
    
     @GetMapping({"/customformat"})
     public ModelAndView getCustomformat()
     {
     	ModelAndView mv = new ModelAndView();
     	
     	mv.setViewName("customformat");
     	return mv;
     }
     @GetMapping({"/dateandtime"})
     public ModelAndView getDateandtime()
     {
    	 ModelAndView mv = new ModelAndView();
    	 
    	 mv.setViewName("dateandtime");
    	 return mv;
     }
     @GetMapping({"/changestyle"})
     public ModelAndView getChangestyle()
     {
    	 ModelAndView mv = new ModelAndView();
    	 
    	 mv.setViewName("changestyle");
    	 return mv;
     }
    @GetMapping({"/index"})
    public ModelAndView getIndexPage()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.setViewName("index");
    	return mv;
    }
    //test
    @GetMapping({"/testlarge"})
    public ModelAndView getTestLarge()
    {
   	 ModelAndView mv = new ModelAndView();
   	 mv.addObject("servletname", "FeatureServlet");
   	 mv.addObject("method", "loadTestLargeFile");
   	 mv.addObject("title", "loadTestLargeFile");
   	 mv.addObject("info", "  load large file for memory test.");
   	 mv.setViewName("general");
   	 return mv;
    }
    //test2
    @GetMapping({"/testlasync"})
    public ModelAndView getTestLAsync()
    {
    	
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "loadLargeFileAsync");
    	mv.addObject("title", "loadLargeFileAsync");
    	mv.addObject("info", "  load large file async for memory test.");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/testload47349"})
    public ModelAndView getTestload47349()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "load47349");
    	mv.addObject("title", "load47349");
    	mv.addObject("info", "  load many image shapes.");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/testload44850"})
    public ModelAndView getTestload44850()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "load44850");
    	mv.addObject("title", "load44850");
    	SimpleDateFormat sdf1 = new SimpleDateFormat("yyyyMMdd");
    	mv.addObject("info", "  文字显示不全."+sdf1.format(new java.util.Date()));
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/testloadSlow"})
    public ModelAndView getTestloadSlow()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "loadSlow");
    	mv.addObject("title", "loadSlow");
    	SimpleDateFormat sdf1 = new SimpleDateFormat("yyyyMMdd");
    	mv.addObject("info", "  卡 变慢了 ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/testloadpwd"})
    public ModelAndView getTestloadPwd()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "loadPwd");
    	mv.addObject("title", "load with password 123456");
    	SimpleDateFormat sdf1 = new SimpleDateFormat("yyyyMMdd");
    	mv.addObject("info", " load with password 123456 不能切换sheet ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/clientjs"})
    public ModelAndView getClientjs()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "testclientjs");
    	mv.addObject("title", "testclientjs");
    	 
    	mv.addObject("info", " testclientjs ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/t45054"})
    public ModelAndView getT45054()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "t45054");
    	mv.addObject("title", "无法切换sheet");
    	 
    	mv.addObject("info", " 无法切换sheet ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/t45229"})
    public ModelAndView getT45229()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "t45229");
    	mv.addObject("title", "又一个文件无法切换sheet");
    	 
    	mv.addObject("info", " 又一个文件无法切换sheet ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/chinesegbcsv"})
    public ModelAndView getChineseCSV()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "chinesegbcsv");
    	mv.addObject("title", "load chinese csv ");
    	 
    	mv.addObject("info", " 中文乱码 ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/t0316file"})
    public ModelAndView get20230316file()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "get20230316file");
    	mv.addObject("title", "1234.xlsx ");
    	 
    	mv.addObject("info", " https://forum.aspose.com/t/23-3-tomcat-demo-xlsx/261639");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/t45251"})
    public ModelAndView get45251file()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "get45251file");
    	mv.addObject("title", "t45251.xls 筛选按钮的位置 确实应该对应移动 ");
    	
    	mv.addObject("info", " https://forum.aspose.com/t/excel/261861/4");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/jpedgeieborder"})
    public ModelAndView getEdgeIEBorder()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "jpedgeieborder");
    	mv.addObject("title", "check border  in edge ie  ");
    	 
    	mv.addObject("info", " 在Aspose.Cells for Java中，使用Edge的IE模式显示Excel表，但在本公司的环境中没有问题，但在客户的环境中不显示Excel表，只输出黑色框。因为在客户环境中也能显示IE，所以我想是Edge的设定问题，如果有什么信息的话请告诉我。 ");
    	mv.setViewName("general");
    	return mv;
    }
    @GetMapping({"/testautofilter"})
    public ModelAndView getTestAutofilter()
    {
    	ModelAndView mv = new ModelAndView();
    	mv.addObject("servletname", "FeatureServlet");
    	mv.addObject("method", "autofilter");
    	mv.addObject("title", "add autofilter");
    	SimpleDateFormat sdf1 = new SimpleDateFormat("yyyyMMdd");
    	mv.addObject("info", "add autofilter ");
    	mv.setViewName("general");
    	return mv;
    }
}