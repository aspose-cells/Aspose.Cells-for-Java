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

package featurescomparison.workingwithcellsrowscolumns.splitpanes.java;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
* How to set split panes
*/
public class ApacheSplitPanes 
{
 public static void main(String[]args) throws Exception 
 {
     
	 String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/splitpanes/data/";
	 
	 Workbook wb = new XSSFWorkbook();
     Sheet sheet = wb.createSheet("new sheet");

     // Create a split with the lower left side being the active quadrant
     sheet.createSplitPane(2000, 2000, 0, 0, Sheet.PANE_LOWER_LEFT);

     FileOutputStream fileOut = new FileOutputStream(dataPath + "ApacheSplitFreezePanes.xlsx");
     wb.write(fileOut);
     fileOut.close();
     System.out.println("Done.");
 }
}