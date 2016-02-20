/* ====================================================================
Licensed to the Apache Software Foundation (ASF) under one or more
contributor license agreements.  See the NOTICE file distributed with
this work for additional information regarding copyright ownership.
The ASF licenses this file to You under the Apache License, Version 2.0
(the "License"); you may not use this file except in compliance with
the License.  You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
==================================================================== */
package featurescomparison.workingwithworkbook.fittoonepage.java;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApacheFitSheetToOnePage {


 public static void main(String[]args) throws Exception 
 {
	 String dataPath = "src/featurescomparison/workingwithworkbook/fittoonepage/data/";
	 
     Workbook wb = new XSSFWorkbook();  //or new HSSFWorkbook();
     Sheet sheet = wb.createSheet("format sheet");
     PrintSetup ps = sheet.getPrintSetup();

     sheet.setAutobreaks(true);

     ps.setFitHeight((short) 1);
     ps.setFitWidth((short) 1);

     // Create various cells and rows for spreadsheet.

     FileOutputStream fileOut = new FileOutputStream(dataPath + "ApacheFitSheetToOnePage.xlsx");
     wb.write(fileOut);
     fileOut.close();
     System.out.println("Done.");
 }
}