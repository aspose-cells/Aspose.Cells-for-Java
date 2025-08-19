package com.aspose.gridjs.demo;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Random;

import javax.net.ssl.HttpsURLConnection;

import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import com.aspose.cells.License;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.gridjs.Config;
import com.aspose.gridjs.GridJsWorkbook;
import com.aspose.gridjs.GridLoadFormat;
import com.aspose.gridjs.demo.test.Util;

import org.springframework.boot.test.context.SpringBootTest;

import com.aspose.gridjs.Config;

@SpringBootTest
public class GridjsdemoApplicationTests {
	private static String getTestfile(String file) {
		 return "D:\\codebase\\local\\gridjstTestdll.Asp.netcore\\files\\"+file;
	}
	private static String getOutputfile(String file) {
		 return "D:\\codebase\\local\\gridjstTestdll.Asp.netcore\\outfiles\\"+file;
	}
	 @BeforeAll
	    public static void init() throws Exception {
		 Config.setFileCacheDirectory("D:/tmpdel/storage/");
		 License lic = new License();
         lic.setLicense("D:\\release\\licesnse\\Aspose.Cells.lic");
		 
	    }
	 @Test
	    public void testConfigEmptySheetMaxRowCol() throws Exception {
		
	        GridJsWorkbook gw = new GridJsWorkbook();
	        Config.setEmptySheetMaxRow(48);
	        Config.setEmptySheetMaxCol(35);
	        Workbook wb = new Workbook();
	        try (java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream()) {
	            wb.save(baos,SaveFormat.XLSX);
	            gw.importExcelFile(new java.io.ByteArrayInputStream(baos.toByteArray()), GridLoadFormat.XLSX);
	            String json = gw.exportToJson();
	            System.out.println(json);
	            org.json.JSONObject cpresult = new org.json.JSONObject(json);
	            
	            String collen = ((org.json.JSONObject)cpresult.getJSONArray("data").get(0)).getJSONObject("cols").getString("len");
	            String rowlen = ((org.json.JSONObject)cpresult.getJSONArray("data").get(0)).getJSONObject("rows").getString("len");
	            assertEquals("36", collen);
	            assertEquals("49", rowlen);
	        }
	    }
	 @Test
	    public void testAutoFillDate() throws Exception {
		    GridJsWorkbook gw = new GridJsWorkbook();
	        gw.importExcelFile(getTestfile("autofill.xlsx"));
	        String json = gw.exportToJson();
	        System.out.println(json);
	        org.json.JSONObject cpresult = new org.json.JSONObject(json);
	        String uid = cpresult.getString("uniqueid");
	        String p = "{\"name\":\"Weekday Schedule\",\"cells\":[],\"src\":{\"sri\":1,\"sci\":2,\"eri\":1,\"eci\":2,\"w\":0,\"h\":0},\"dest\":{\"sri\":2,\"sci\":2,\"eri\":7,\"eci\":2,\"w\":0,\"h\":0},\"fromsheet\":\"Weekday Schedule\",\"what\":\"all\",\"op\":\"batchupdate\"}";
	        String ret = gw.updateCell(p, uid);
	        System.out.println(ret);
	        String expect = "{\"op\":\"update\",\"data\":[{\"name\":\"Weekday Schedule\",\"cells\":[{\"ri\":2,\"ci\":2,\"text\":\"28-Feb\",\"ufv\":\"2021-2-28\",\"dt\":\"20210228 000000\"},{\"ri\":3,\"ci\":2,\"text\":\"1-Mar\",\"ufv\":\"2021-3-1\",\"dt\":\"20210301 000000\"},{\"ri\":4,\"ci\":2,\"text\":\"2-Mar\",\"ufv\":\"2021-3-2\",\"dt\":\"20210302 000000\"},{\"ri\":5,\"ci\":2,\"text\":\"3-Mar\",\"ufv\":\"2021-3-3\",\"dt\":\"20210303 000000\"},{\"ri\":6,\"ci\":2,\"text\":\"4-Mar\",\"ufv\":\"2021-3-4\",\"dt\":\"20210304 000000\"},{\"ri\":7,\"ci\":2,\"text\":\"5-Mar\",\"ufv\":\"2021-3-5\",\"dt\":\"20210305 000000\"}]}]}";
	        assertEquals(expect, ret);
	    }

	 @Test
	    public void testAutoFillFormula() throws Exception {
		   GridJsWorkbook gw = new GridJsWorkbook();
	        String path =  getTestfile("autofill.xlsx");
	        try (java.io.FileInputStream fs = new java.io.FileInputStream(path)) {
	            gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(java.nio.file.Paths.get(path).getFileName().toString()), null);
	        }
	        
	        String json = gw.exportToJson();
	        System.out.println(json);
	        org.json.JSONObject cpresult = new org.json.JSONObject(json);
	        String uid = cpresult.getString("uniqueid");
	        String p = "{\"name\":\"Weekday Schedule\",\"cells\":[{\"ri\":18,\"ci\":0,\"text\":\"3\"},{\"ri\":20,\"ci\":0,\"text\":\"3\"},{\"ri\":22,\"ci\":0,\"text\":\"3\"},{\"ri\":24,\"ci\":0,\"text\":\"3\"},{\"ri\":26,\"ci\":0,\"text\":\"3\"},{\"ri\":18,\"ci\":1,\"text\":\"5.6\"},{\"ri\":20,\"ci\":1,\"text\":\"5.6\"},{\"ri\":22,\"ci\":1,\"text\":\"5.6\"},{\"ri\":24,\"ci\":1,\"text\":\"5.6\"},{\"ri\":26,\"ci\":1,\"text\":\"5.6\"},{\"ri\":19,\"ci\":0,\"text\":\"4\"},{\"ri\":21,\"ci\":0,\"text\":\"4\"},{\"ri\":23,\"ci\":0,\"text\":\"4\"},{\"ri\":25,\"ci\":0,\"text\":\"4\"},{\"ri\":27,\"ci\":0,\"text\":\"4\"},{\"ri\":19,\"ci\":1,\"text\":\"6.6\"},{\"ri\":21,\"ci\":1,\"text\":\"6.6\"},{\"ri\":23,\"ci\":1,\"text\":\"6.6\"},{\"ri\":25,\"ci\":1,\"text\":\"6.6\"},{\"ri\":27,\"ci\":1,\"text\":\"6.6\"}],\"src\":{\"sri\":16,\"sci\":0,\"eri\":17,\"eci\":2,\"w\":0,\"h\":0},\"dest\":{\"sri\":18,\"sci\":0,\"eri\":27,\"eci\":2,\"w\":0,\"h\":0},\"what\":\"all\",\"op\":\"batchupdate\"}";
	        String ret = gw.updateCell(p, uid);
	        System.out.println(ret);
	        String expect = "{\"op\":\"update\",\"data\":[{\"name\":\"Weekday Schedule\",\"cells\":[{\"ri\":18,\"ci\":0,\"text\":\"3\"},{\"ri\":18,\"ci\":1,\"text\":\"5.6\"},{\"ri\":18,\"ci\":2,\"text\":\"8.6\",\"f\":\"=SUM(A19:B19)\"},{\"ri\":19,\"ci\":0,\"text\":\"4\"},{\"ri\":19,\"ci\":1,\"text\":\"6.6\"},{\"ri\":19,\"ci\":2,\"text\":\"2.841470985\",\"f\":\"=2+SIN(1)\"},{\"ri\":20,\"ci\":0,\"text\":\"3\"},{\"ri\":20,\"ci\":1,\"text\":\"5.6\"},{\"ri\":20,\"ci\":2,\"text\":\"8.6\",\"f\":\"=SUM(A21:B21)\"},{\"ri\":21,\"ci\":0,\"text\":\"4\"},{\"ri\":21,\"ci\":1,\"text\":\"6.6\"},{\"ri\":21,\"ci\":2,\"text\":\"2.841470985\",\"f\":\"=2+SIN(1)\"},{\"ri\":22,\"ci\":0,\"text\":\"3\"},{\"ri\":22,\"ci\":1,\"text\":\"5.6\"},{\"ri\":22,\"ci\":2,\"text\":\"8.6\",\"f\":\"=SUM(A23:B23)\"},{\"ri\":23,\"ci\":0,\"text\":\"4\"},{\"ri\":23,\"ci\":1,\"text\":\"6.6\"},{\"ri\":23,\"ci\":2,\"text\":\"2.841470985\",\"f\":\"=2+SIN(1)\"},{\"ri\":24,\"ci\":0,\"text\":\"3\"},{\"ri\":24,\"ci\":1,\"text\":\"5.6\"},{\"ri\":24,\"ci\":2,\"text\":\"8.6\",\"f\":\"=SUM(A25:B25)\"},{\"ri\":25,\"ci\":0,\"text\":\"4\"},{\"ri\":25,\"ci\":1,\"text\":\"6.6\"},{\"ri\":25,\"ci\":2,\"text\":\"2.841470985\",\"f\":\"=2+SIN(1)\"},{\"ri\":26,\"ci\":0,\"text\":\"3\"},{\"ri\":26,\"ci\":1,\"text\":\"5.6\"},{\"ri\":26,\"ci\":2,\"text\":\"8.6\",\"f\":\"=SUM(A27:B27)\"},{\"ri\":27,\"ci\":0,\"text\":\"4\"},{\"ri\":27,\"ci\":1,\"text\":\"6.6\"},{\"ri\":27,\"ci\":2,\"text\":\"2.841470985\",\"f\":\"=2+SIN(1)\"}]}]}";
	        assertEquals(expect, ret);
	    }
	 
		@Test
		public void testShapeRotation() throws Exception {
			GridJsWorkbook gw = new GridJsWorkbook();

			GridJsWorkbook.CacheImp = new LocalFileCache();
			String path = getTestfile("pictest.xls");
			// Assuming there's a method to export to JSON
			try (FileInputStream fs = new FileInputStream(path)) {
				gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(path), null);
			}
			String json = gw.exportToJson();

			// Parse JSON and manipulate the workbook
			JSONObject jsonObject = new JSONObject(json);
			String uniqueId = jsonObject.getString("uniqueid");

			String p = "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[{\"left\":735.7283738306762,\"top\":124.21221762187565,\"originAngle\":14,\"angle\":62.255290286440825,\"zorder\":15,\"width\":189,\"height\":46,\"id\":\"6\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";
			gw.mergeExcelFileFromJson(uniqueId, p);

			// Save the modified workbook to a MemoryStream equivalent in Java
			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			gw.saveToExcelFile(outputStream);

			// Reset the position of the output stream to the beginning
			ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
			Workbook workbook = new Workbook(inputStream);

			// Access the shape and get its rotation angle
			Worksheet sheet = workbook.getWorksheets().get("Sheet2");
			Shape shape = (Shape) sheet.getShapes().get(6); // Assuming the shape is the first child
			double rotation = shape.getRotationAngle();

			// Assert the expected rotation value
			assertEquals(62.0, rotation, 0.001);
		}
	 
		@Test
		public void testImageRotation() throws Exception {
			GridJsWorkbook gw = new GridJsWorkbook();

			GridJsWorkbook.CacheImp = new LocalFileCache();
			String path = getTestfile("pictest.xls");
			// Assuming there's a method to export to JSON
			try (FileInputStream fs = new FileInputStream(path)) {
				gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(path), null);
			}
			String json = gw.exportToJson();

			// Parse JSON and manipulate the workbook
			JSONObject jsonObject = new JSONObject(json);
			String uniqueId = jsonObject.getString("uniqueid");

			String p = "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":2}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[{\"left\":486,\"top\":428.9999999999997,\"originAngle\":18.2446899414063,\"angle\":229.7013187480324,\"zorder\":15,\"width\":206,\"height\":71,\"id\":\"2\"}],\"shapes\":[]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";
			gw.mergeExcelFileFromJson(uniqueId, p);

			// Save the modified workbook to a MemoryStream equivalent in Java
			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			gw.saveToExcelFile(outputStream);

			// Reset the position of the output stream to the beginning
			ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
			Workbook workbook = new Workbook(inputStream);

			// Access the shape and get its rotation angle
			Worksheet sheet = workbook.getWorksheets().get("Sheet2");
			Shape shape = (Shape) sheet.getPictures().get(2); // Assuming the shape is the first child
			double rotation = shape.getRotationAngle();

			// Assert the expected rotation value
			assertEquals(-130, rotation, 1);
		}
		
		@Test
		public void testShapeMove() throws Exception {
			GridJsWorkbook gw = new GridJsWorkbook();

			GridJsWorkbook.CacheImp = new LocalFileCache();
			String path = getTestfile("pictest.xls");
			// Assuming there's a method to export to JSON
			try (FileInputStream fs = new FileInputStream(path)) {
				gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(path), null);
			}
			String json = gw.exportToJson();

			// Parse JSON and manipulate the workbook
			JSONObject jsonObject = new JSONObject(json);
			String uniqueId = jsonObject.getString("uniqueid");

			String p =  "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":2}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[{\"left\":86.13666886109286,\"top\":38,\"originAngle\":180,\"angle\":180,\"zorder\":15,\"width\":189,\"height\":46,\"id\":\"7\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";
			gw.mergeExcelFileFromJson(uniqueId, p);

			// Save the modified workbook to a MemoryStream equivalent in Java
			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			gw.saveToExcelFile(outputStream);

			// Reset the position of the output stream to the beginning
			ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
			Workbook workbook = new Workbook(inputStream);
            int shapeid = 7;
			// Access the shape and get its rotation angle
			Worksheet sheet = workbook.getWorksheets().get("Sheet2");
			Shape shape = (Shape) sheet.getShapes().get(shapeid); // Assuming the shape is the first child
            int left = shape.getLeftToCorner();
            int top = shape.getTopToCorner();

			// Assert the expected rotation value
			assertEquals(86, left);
			assertEquals(38, top);
		}
		
		@Test
		public void testShapeResize() throws Exception {
			GridJsWorkbook gw = new GridJsWorkbook();

			GridJsWorkbook.CacheImp = new LocalFileCache();
			String path = getTestfile("pictest.xls");
			// Assuming there's a method to export to JSON
			try (FileInputStream fs = new FileInputStream(path)) {
				gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(path), null);
			}
			String json = gw.exportToJson();

			// Parse JSON and manipulate the workbook
			JSONObject jsonObject = new JSONObject(json);
			String uniqueId = jsonObject.getString("uniqueid");

			String p =  "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":2}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[{\"left\":84.61946061188769,\"top\":223.9999999999995,\"originAngle\":180,\"angle\":155.2060739658465,\"zorder\":18,\"width\":238.96908007635514,\"height\":58.161786685250455,\"id\":\"7\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";
			gw.mergeExcelFileFromJson(uniqueId, p);

			// Save the modified workbook to a MemoryStream equivalent in Java
			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			gw.saveToExcelFile(outputStream);

			// Reset the position of the output stream to the beginning
			ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
			Workbook workbook = new Workbook(inputStream);
            int shapeid = 7;
			// Access the shape and get its rotation angle
			Worksheet sheet = workbook.getWorksheets().get("Sheet2");
			Shape shape = (Shape) sheet.getShapes().get(shapeid); // Assuming the shape is the first child
            int w = shape.getWidth();
            int h = shape.getHeight();

			// Assert the expected rotation value
			
			assertEquals(58, h);
			//238 or 239 it not very accurate
			assertEquals(238, w);
		}
		
		@Test
		public void testImageMove() throws Exception {
			GridJsWorkbook gw = new GridJsWorkbook();

			GridJsWorkbook.CacheImp = new LocalFileCache();
			String path = getTestfile("pictest.xls");
			// Assuming there's a method to export to JSON
			try (FileInputStream fs = new FileInputStream(path)) {
				gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(path), null);
			}
			String json = gw.exportToJson();

			// Parse JSON and manipulate the workbook
			JSONObject jsonObject = new JSONObject(json);
			String uniqueId = jsonObject.getString("uniqueid");

			String p =  "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":2}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[{\"left\":296.9292955892031,\"top\":75.13595091372946,\"originAngle\":326.733581542969,\"angle\":8.434664397787628,\"zorder\":24,\"width\":334,\"height\":201,\"id\":\"0\"}],\"shapes\":[{\"left\":84.61946061188769,\"top\":223.9999999999995,\"originAngle\":180,\"angle\":155.2060739658465,\"zorder\":18,\"width\":238.96908007635514,\"height\":58.161786685250455,\"id\":\"7\"},{\"left\":803.0758393680053,\"top\":58,\"originAngle\":0,\"angle\":0,\"zorder\":22,\"width\":375,\"height\":128,\"id\":\"13\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";
			gw.mergeExcelFileFromJson(uniqueId, p);

			// Save the modified workbook to a MemoryStream equivalent in Java
			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			gw.saveToExcelFile(outputStream);

			// Reset the position of the output stream to the beginning
			ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
			Workbook workbook = new Workbook(inputStream);
            int shapeid = 0;
			// Access the shape and get its rotation angle
			Worksheet sheet = workbook.getWorksheets().get("Sheet2");
			Shape shape = (Shape) sheet.getPictures().get(shapeid); // Assuming the shape is the first child
            int left = shape.getLeftToCorner();
            int top = shape.getTopToCorner();

			// Assert the expected rotation value
            //297->296
			assertEquals(296, left);
			assertEquals(75, top);
		}
		@Test
		public void testImageResize() throws Exception {
			GridJsWorkbook gw = new GridJsWorkbook();

			GridJsWorkbook.CacheImp = new LocalFileCache();
			String path = getTestfile("pictest.xls");
			// Assuming there's a method to export to JSON
			try (FileInputStream fs = new FileInputStream(path)) {
				gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(path), null);
			}
			String json = gw.exportToJson();

			// Parse JSON and manipulate the workbook
			JSONObject jsonObject = new JSONObject(json);
			String uniqueId = jsonObject.getString("uniqueid");

			String p =  "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":1}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[{\"left\":198.22518261373216,\"top\":282.4330962241735,\"originAngle\":326.733581542969,\"angle\":351.3970493213394,\"zorder\":27,\"width\":211.5761326868053,\"height\":127.32575649714929,\"id\":\"0\"}],\"shapes\":[{\"left\":84.61946061188769,\"top\":223.9999999999995,\"originAngle\":180,\"angle\":155.2060739658465,\"zorder\":18,\"width\":238.96908007635514,\"height\":58.161786685250455,\"id\":\"7\"},{\"left\":391.23857801185045,\"top\":209.49026704645462,\"originAngle\":0,\"angle\":0,\"zorder\":30,\"width\":292.1794562603063,\"height\":99.73058773685122,\"id\":\"13\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";
			gw.mergeExcelFileFromJson(uniqueId, p);

			// Save the modified workbook to a MemoryStream equivalent in Java
			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			gw.saveToExcelFile(outputStream);

			// Reset the position of the output stream to the beginning
			ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
			Workbook workbook = new Workbook(inputStream);
            int shapeid = 0;
			// Access the shape and get its rotation angle
			Worksheet sheet = workbook.getWorksheets().get("Sheet2");
			Shape shape = (Shape) sheet.getPictures().get(shapeid); // Assuming the shape is the first child
            int w = shape.getWidth();
            int h = shape.getHeight();

			// Assert the expected rotation value
			
			assertEquals(127, h);
			//212 ->211 not very accurate
			assertEquals(211, w);
		}
		//now the temp cache file is xlsx format ,so the below test will not raise insert fail again
		@Test
		public void testInsertRowFail() throws Exception {
			GridJsWorkbook gw = new GridJsWorkbook();

			GridJsWorkbook.CacheImp = new LocalFileCache();
			String path = getTestfile("addrowfail.xls");
			// Assuming there's a method to export to JSON
			try (FileInputStream fs = new FileInputStream(path)) {
				gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(path), null);
			}
			String json = gw.exportToJson();

			// Parse JSON and manipulate the workbook
			JSONObject jsonObject = new JSONObject(json);
			String uniqueId = jsonObject.getString("uniqueid");

			String p =  "{\"name\":\"Worksheet\",\"ri\":1,\"n\":1,\"h\":12,\"op\":\"insertrow\"}";
			String ret=gw.updateCell( p,uniqueId);

			 String expect = "{\"op\":\"update\",\"data\":[]}";

			 assertEquals(expect, ret);
		}
		@Test
		 public void testConfigSkipShapes() throws Exception {
		        Config.setSkipInvisibleShapes(false);
		        checkShapeCount(6);
		        //default
		        Config.setSkipInvisibleShapes(true);
		        checkShapeCount(3);
		    }

		    private static void checkShapeCount(int expectedCount) throws Exception {
		    	GridJsWorkbook gw = new GridJsWorkbook();
		    	GridJsWorkbook.CacheImp = new LocalFileCache();
		        String path = getTestfile("testfileview gridjs.xlsx"); // Assuming util.getTestfile is a static method in this context
		        gw.importExcelFile(path);

		        ByteArrayOutputStream baos = new ByteArrayOutputStream();
		        gw.jsonToStream(baos, "hello2.xlsx");

		        baos.close();

		        String json = new String(baos.toByteArray(), StandardCharsets.UTF_8);
		        JSONObject cpresult = new JSONObject(json);
		        JSONArray shapesArray = ((JSONArray) ((JSONArray) cpresult.getJSONArray("data")).getJSONObject(0).getJSONArray("shapes"));
		        int len = shapesArray.length();
		        assertEquals(expectedCount, len);
		    }
		   
		    @Test
		    public void testInsertImageWithNoStreamCacheImp() throws Exception {
		    	GridJsWorkbook gw = new GridJsWorkbook();
		        gw.CacheImp=(null); // Ensure the cache implementation is null for this test

		        String path = getTestfile("pictest.xls");
		        gw.importExcelFile(path);

		        String json = gw.exportToJson();
		        System.out.println(json);
		        JSONObject cpresult = new JSONObject(json);
		        String uid = cpresult.getString("uniqueid");
		        String p = "{\"name\":\"Sheet2\",\"ri\":1,\"ci\":1}";
		        String picRet = null;

		        File imageFile = new File(getTestfile("snap1.png"));
		        try (FileInputStream s = new FileInputStream(imageFile)) {
		            picRet = gw.insertImage(uid, p, s, null);
		        }
		        System.out.println("insert first image : " + picRet);
		        JSONObject picRetJson = new JSONObject(picRet);

		        int picId = picRetJson.getInt("id");
		        assertEquals(3, picId);

		        ByteArrayOutputStream baos = new ByteArrayOutputStream();
		        gw.saveToExcelFile(baos);

		        baos.close();

		        byte[] byteArray = baos.toByteArray();
		        try (ByteArrayInputStream bais = new ByteArrayInputStream(byteArray)) {
		            Workbook wb = new Workbook(bais);

		            Worksheet sheet = wb.getWorksheets().get("Sheet2"); // Assuming "Sheet2" is the second sheet
		            Shape sp = (Shape) sheet.getPictures().get(picId); // Shapes are typically stored in a list of sheet objects, and the ID might need to be adjusted

		            int w = sp.getWidth();
		            int h = sp.getHeight();

		            String expectW = "320";
		            String expectH = "204";
		            assertEquals(expectW, Integer.toString(w));
		            assertEquals(expectH, Integer.toString(h));
		            assertEquals(expectW, picRetJson.getString("width"));
		            assertEquals(expectH, picRetJson.getString("height"));
		            System.out.println("pic count:"+sheet.getPictures().getCount());

		            // Add second image with a new GridJsWorkbook instance
		            
		            p = "{\"name\":\"Sheet2\",\"ri\":3,\"ci\":3}";
		            try (FileInputStream s = new FileInputStream(imageFile)) {
		                picRet = gw.insertImage(uid, p, s, null);
		                System.out.println("insert second image : " + picRet);
		            }
		            picRetJson = new JSONObject(picRet);

		            baos = new ByteArrayOutputStream();
		            gw.saveToExcelFile(baos);

		            baos.close();

		            byteArray = baos.toByteArray();
		            try (ByteArrayInputStream bais2 = new ByteArrayInputStream(byteArray)) {
		                Workbook wb2 = new Workbook(bais2);
		                picId = picRetJson.getInt("id");
		                String urlString = picRetJson.getString("url");
		                assertTrue(urlString.indexOf("uid=") > 0 && urlString.indexOf("&id=") > 0);
		                assertEquals(4, picId);

		                sheet = wb2.getWorksheets().get("Sheet2"); 
		                
		                sp = (Shape) sheet.getPictures().get(picId);
		                w = sp.getWidth();
		                h = sp.getHeight();
		                assertEquals(expectW, Integer.toString(w));
		                assertEquals(expectH, Integer.toString(h));

		                sp = (Shape) sheet.getPictures().get(3);  
		                w = sp.getWidth();
		                h = sp.getHeight();
		                assertEquals(expectW, Integer.toString(w));
		                assertEquals(expectH, Integer.toString(h));
		            }
		        }
		    }
		    
		    @Test
            public void TestInsertImageWithStreamCacheImp()
            {
               String[] uid = new String[1];
                try {
					insertImage(null,  uid);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

            } 
		    
		    @Test
		    public void testOptimizeForSameImage() throws Exception {
		        // Do not limit shape or image count
		        Config.setIslimitShapeOrImage(false);
		        Config.setSameImageDetecting(true);
		        

		        GridJsWorkbook gw = new GridJsWorkbook();
		    	GridJsWorkbook.CacheImp = new LocalFileCache();
		        Random rd = new Random();
		        Workbook wb = new Workbook();

		        // Add only one picture to many cells
		        addPictureToWb(rd, wb, getTestfile("snap1.png"));

		        ByteArrayOutputStream baos = new ByteArrayOutputStream();
		        wb.save(baos,SaveFormat.XLSX);
                
		        baos.close();

		        byte[] byteArray = baos.toByteArray();
		        try (ByteArrayInputStream bais = new ByteArrayInputStream(byteArray)) {
		            gw.importExcelFile(bais, GridLoadFormat.XLSX, null);
		        }

		        String json = gw.exportToJson();

		        JSONObject cpresult = new JSONObject(json);
		        String uid = cpresult.getString("uniqueid");
		        // Only one picture in zip
		        testImageZipFiles(uid, 2);

		        // Add more pictures
		        addPictureToWb(rd, wb, getTestfile("pic2.png"));
		        addPictureToWb(rd, wb, getTestfile("pic3.png"));

		        baos = new ByteArrayOutputStream();
		        wb.save(baos,SaveFormat.XLSX);

		        baos.close();

		        byteArray = baos.toByteArray();
		        try (ByteArrayInputStream bais = new ByteArrayInputStream(byteArray)) {
		            gw.importExcelFile(bais, GridLoadFormat.XLSX, null);
		        }

		        long startTime = System.currentTimeMillis();
		        json = gw.exportToJson();
		        long endTime = System.currentTimeMillis();

		        long costTime1 = endTime - startTime;
		        cpresult = new JSONObject(json);
		        uid = cpresult.getString("uniqueid");

		        // Only 3 images actually
		        testImageZipFiles(uid, 4);

		        // No same image detection
		        Config.setSameImageDetecting(false);

		        baos = new ByteArrayOutputStream();
		        wb.save(baos,SaveFormat.XLSX);

		        baos.close();

		        byteArray = baos.toByteArray();
		        try (ByteArrayInputStream bais = new ByteArrayInputStream(byteArray)) {
		            gw.importExcelFile(bais, GridLoadFormat.XLSX, null);
		        }

		        startTime = System.currentTimeMillis();
		        json = gw.exportToJson();
		        endTime = System.currentTimeMillis();

		        long costTime2 = endTime - startTime;
		        cpresult = new JSONObject(json);
		        uid = cpresult.getString("uniqueid");

		        // Total image is much more
		        testImageZipFiles(uid, 693);

		        // Less I/O, less time
		        assertEquals(true, costTime2 > costTime1);

		        // Restore Config to default value after test
		        Config.setIslimitShapeOrImage(true);
		        Config.setSameImageDetecting(true);
		    }
		    
		    @Test
		    public void testInsertImageAndThenCopyImage() throws Exception {
		        String[] uidarray = new String[1];
		        // Insert image and the picid is 3
		        insertImage(null, uidarray);
		        String uid=uidarray[0];
		        // srcid set to 3
		        String p = "{\"name\":\"Sheet2\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":3,\"srcid\":149,\"isshape\":false}";
		        String expectW = "320";
		        String expectH = "204";
		        String expectErrMsg = "{\"Error\":\"wrong picture id \"}";

		        GridJsWorkbook gw = new GridJsWorkbook();
		        GridJsWorkbook.CacheImp = new LocalFileCache();
		        String picRet = gw.copyImageOrShape(uid, p);

		        System.out.println("Copy image : " + picRet);
		        JSONObject picRetJson = new JSONObject(picRet);

		        // Copy image return id shall equal with count
		        int picId = picRetJson.getInt("id");
		        assertEquals(5, picId);

		        String picUrl = picRetJson.getString("url");

		        int idIdx = picUrl.indexOf("id=");
		        String downloadId = picUrl.substring(idIdx + 3);
		        System.out.println(downloadId);

		        // Image not existed
		        // InputStream fsReader = gw.getCacheImp().loadStream(downloadId);
		        // assertEquals(true, fsReader.available() > 0);

		        ByteArrayOutputStream baos = new ByteArrayOutputStream();
		        gw.saveToExcelFile(baos);
		        baos.close();

		        byte[] byteArray = baos.toByteArray();
		        try (ByteArrayInputStream bais = new ByteArrayInputStream(byteArray)) {
		            Workbook wb = new Workbook(bais);

		            Shape sp = wb.getWorksheets().get("Sheet2").getPictures().get(picId);

		            int w = sp.getWidth();
		            int h = sp.getHeight();

		            assertEquals(expectW, Integer.toString(w));
		            assertEquals(expectH, Integer.toString(h));
		            assertEquals(expectW, picRetJson.getString("width"));
		            assertEquals(expectH, picRetJson.getString("height"));
		            assertEquals(11, sp.getUpperLeftRow());
		            assertEquals(3, sp.getUpperLeftColumn());
		        }
		    }
		    @Test
		    public void testInsertImageWithStreamCacheImpWithURL()
	        { String[] uid = new String[1];
	            try {
					insertImage("https://www.baidu.com/img/flexible/logo/pc/result.png",   uid);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

	        }
		    @Test
		    public void testDeleteImageWithStreamCacheImp() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp=(new LocalFileCache());

		        // Load the workbook using Aspose.Cells
		        String testFilePath = getTestfile("pictest.xls");
		        Workbook workbook = new Workbook(new FileInputStream(testFilePath));
		        Worksheet sheet = workbook.getWorksheets().get("Sheet2");
		        int originCount = sheet.getPictures().getCount();

		        // Import the Excel file using the custom GridJsWorkbook
		        try (FileInputStream fs = new FileInputStream(testFilePath)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), "123");
		        }

		        // Export to JSON and manipulate the JSON string
		        String json = gridJsWorkbook.exportToJson();
		        System.out.println(json);
		        JSONObject jsonObject = new JSONObject(json);
		        String uid = jsonObject.getString("uniqueid");

		        // Prepare the JSON string 'p' as per original C# code
		        String p =  "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":2}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[{\"left\":118,\"top\":270.0000000000001,\"originAngle\":326.733581542969,\"angle\":353.07958890754185,\"zorder\":23,\"width\":334,\"height\":201,\"id\":\"0\"},{\"op\":\"del\",\"id\":\"1\"},{\"left\":486,\"top\":429,\"originAngle\":18.2446899414063,\"angle\":18.2446899414063,\"zorder\":21,\"width\":206,\"height\":71,\"id\":\"2\"}],\"shapes\":[{\"left\":735.7283738306762,\"top\":124.21221762187571,\"originAngle\":14,\"angle\":316.43181320984263,\"zorder\":18,\"width\":189,\"height\":46,\"id\":\"6\"},{\"left\":676.1465596207154,\"top\":230.79871381498958,\"originAngle\":348,\"angle\":26.971249728296755,\"zorder\":16,\"width\":189,\"height\":46,\"id\":\"9\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]} "; // You need to construct this JSON string as per your needs

		        // Merge changes from JSON
		        gridJsWorkbook.mergeExcelFileFromJson(uid, p);

		        // Save the workbook to a ByteArrayOutputStream
		        ByteArrayOutputStream ms = new ByteArrayOutputStream();
		        gridJsWorkbook.saveToExcelFile(ms);
		        ms.flush();
		        ByteArrayInputStream inputStream = new ByteArrayInputStream(ms.toByteArray());
		        workbook = new Workbook(inputStream);

		        // Check the count of pictures after merging JSON changes
		        sheet = workbook.getWorksheets().get("Sheet2");
		        int count = sheet.getPictures().getCount();
		        assertEquals(originCount - 1, count);

		        // Mock second post delete 2 pictures
		        // Repeat the process with a new JSON string that includes the delete operations
		        // for pictures with IDs "1" and "2".
		        p = "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":2}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[{\"left\":118,\"top\":270.0000000000001,\"originAngle\":326.733581542969,\"angle\":353.07958890754185,\"zorder\":23,\"width\":334,\"height\":201,\"id\":\"0\"},{\"op\":\"del\",\"id\":\"1\"},{ \"op\":\"del\",\"id\":\"2\"}],\"shapes\":[{\"left\":735.7283738306762,\"top\":124.21221762187571,\"originAngle\":14,\"angle\":316.43181320984263,\"zorder\":18,\"width\":189,\"height\":46,\"id\":\"6\"},{\"left\":676.1465596207154,\"top\":230.79871381498958,\"originAngle\":348,\"angle\":26.971249728296755,\"zorder\":16,\"width\":189,\"height\":46,\"id\":\"9\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}"; // Construct the new JSON string for the second test

		        gridJsWorkbook = new GridJsWorkbook();
		        
		        
		        gridJsWorkbook.mergeExcelFileFromJson(uid, p);
		        ms = new ByteArrayOutputStream();
		        gridJsWorkbook.saveToExcelFile(ms);
		        inputStream = new ByteArrayInputStream(ms.toByteArray());
		        workbook = new Workbook(inputStream);
		        sheet = workbook.getWorksheets().get("Sheet2");
		        count = sheet.getPictures().getCount();
		        assertEquals(originCount - 2, count);
		    }
		    
		    @Test
		    public void testDeleteShapeWithStreamCacheImp() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp=(new LocalFileCache());

		        // Load the workbook using Aspose.Cells
		        String testFilePath = getTestfile("pictest.xls");
		        Workbook workbook = new Workbook(new FileInputStream(testFilePath));
		        Worksheet sheet = workbook.getWorksheets().get("Sheet2");
		        int originCount = sheet.getShapes().getCount();

		        // Import the Excel file using the custom GridJsWorkbook
		        try (FileInputStream fs = new FileInputStream(testFilePath)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), "123");
		        }

		        // Export to JSON and manipulate the JSON string
		        String json = gridJsWorkbook.exportToJson();
		        System.out.println(json);
		        JSONObject jsonObject = new JSONObject(json);
		        String uid = jsonObject.getString("uniqueid");

		        // Prepare the JSON string 'p' as per original C# code
		        String p =  " {\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":1}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[{\"left\":221.15803068140252,\"top\":92.19221273648674,\"originAngle\":80,\"angle\":80,\"zorder\":18,\"width\":134,\"height\":48,\"id\":\"0\"},{\"left\":265.9518104015799,\"top\":229.88239564428312,\"originAngle\":0,\"angle\":0,\"zorder\":16,\"width\":138,\"height\":46,\"id\":\"1\"},{\"left\":281.5786750570451,\"top\":102.38889291739602,\"originAngle\":120,\"angle\":120,\"zorder\":17,\"width\":134,\"height\":48,\"id\":\"3\"}, {\"left\":370.86484887511415,\"top\":212.08212536113388,\"originAngle\":170,\"angle\":170,\"zorder\":20,\"width\":189,\"height\":46,\"id\":\"5\"},{\"op\":\"del\" ,\"id\":\"6\"},{\"left\":495.97472021066494,\"top\":217.8867513611616,\"originAngle\":180,\"angle\":180,\"zorder\":21,\"width\":189,\"height\":46,\"id\":\"7\"},{\"left\":567.1896142619267,\"top\":311.71361159049684,\"originAngle\":348,\"angle\":348,\"zorder\":22,\"width\":189,\"height\":46,\"id\":\"9\"},{\"left\":724.3627247905652,\"top\":338.6218715229195,\"originAngle\":299.491744995117,\"angle\":338.41234009876024,\"zorder\":24,\"width\":182,\"height\":48,\"id\":\"10\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";

		        // Merge changes from JSON
		        gridJsWorkbook.mergeExcelFileFromJson(uid, p);

		        // Save the workbook to a ByteArrayOutputStream
		        ByteArrayOutputStream ms = new ByteArrayOutputStream();
		        gridJsWorkbook.saveToExcelFile(ms);
		        ms.flush();
		        ByteArrayInputStream inputStream = new ByteArrayInputStream(ms.toByteArray());
		        workbook = new Workbook(inputStream);

		        // Check the count of pictures after merging JSON changes
		        sheet = workbook.getWorksheets().get("Sheet2");
		        int count = sheet.getShapes().getCount();
		        assertEquals(originCount - 1, count);

		        // Mock second post delete 2 pictures
		        // Repeat the process with a new JSON string that includes the delete operations
		        // for pictures with IDs "1" and "2".
		        p = "{\"sheetname\":\"Sheet2\",\"actrow\":21,\"actcol\":9,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":25},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"Black\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"Black\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":true}}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72},\"16\":{\"width\":72},\"17\":{\"width\":72},\"18\":{\"width\":72},\"19\":{\"width\":72},\"20\":{\"width\":72}},\"rows\":{\"21\":{\"height\":21,\"cells\":{\"9\":{\"text\":\"\",\"style\":1}}},\"height\":20,\"len\":27},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[{\"left\":221.15803068140252,\"top\":92.19221273648674,\"originAngle\":80,\"angle\":80,\"zorder\":18,\"width\":134,\"height\":48,\"id\":\"0\"},{\"left\":265.9518104015799,\"top\":229.88239564428312,\"originAngle\":0,\"angle\":0,\"zorder\":16,\"width\":138,\"height\":46,\"id\":\"1\"},{\"left\":281.5786750570451,\"top\":102.38889291739602,\"originAngle\":120,\"angle\":120,\"zorder\":17,\"width\":134,\"height\":48,\"id\":\"3\"},{\"op\":\"del\" ,\"id\":\"4\"},{\"left\":370.86484887511415,\"top\":212.08212536113388,\"originAngle\":170,\"angle\":170,\"zorder\":20,\"width\":189,\"height\":46,\"id\":\"5\"},{\"op\":\"del\" ,\"id\":\"6\"},{\"left\":495.97472021066494,\"top\":217.8867513611616,\"originAngle\":180,\"angle\":180,\"zorder\":21,\"width\":189,\"height\":46,\"id\":\"7\"},{\"left\":567.1896142619267,\"top\":311.71361159049684,\"originAngle\":348,\"angle\":348,\"zorder\":22,\"width\":189,\"height\":46,\"id\":\"9\"},{\"left\":724.3627247905652,\"top\":338.6218715229195,\"originAngle\":299.491744995117,\"angle\":338.41234009876024,\"zorder\":24,\"width\":182,\"height\":48,\"id\":\"10\"}]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{\"0\":{\"width\":72},\"1\":{\"width\":72},\"2\":{\"width\":72},\"3\":{\"width\":72},\"4\":{\"width\":72},\"5\":{\"width\":72},\"6\":{\"width\":72},\"7\":{\"width\":72},\"8\":{\"width\":72},\"9\":{\"width\":72},\"10\":{\"width\":72},\"11\":{\"width\":72},\"12\":{\"width\":72},\"13\":{\"width\":72},\"14\":{\"width\":72},\"15\":{\"width\":72}},\"rows\":{\"height\":20,\"len\":20},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";

		        gridJsWorkbook = new GridJsWorkbook();
		        
		        
		        gridJsWorkbook.mergeExcelFileFromJson(uid, p);
		        ms = new ByteArrayOutputStream();
		        gridJsWorkbook.saveToExcelFile(ms);
		        inputStream = new ByteArrayInputStream(ms.toByteArray());
		        workbook = new Workbook(inputStream);
		        sheet = workbook.getWorksheets().get("Sheet2");
		        count = sheet.getShapes().getCount();
		        assertEquals(originCount - 2, count);
		    }
		    
		    @Test
		    public void testCELLSAPP1039() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp=(new LocalFileCache());

		        // Import the Excel file using the custom GridJsWorkbook
		        String path = getTestfile("CELLSAPP-1039.xlsx");
		        try (FileInputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xlsx"), null);
		        }

		        // Set configuration limits
		        Config.setMaxShapeOrImageCount(5);
		        Config.setMaxTotalShapeOrImageCount(100);
		        long startTime = System.currentTimeMillis();
		        System.out.println("TestCELLSAPP1039 " + Config.getIslimitShapeOrImage());

		        // Export to JSON
		        String json = gridJsWorkbook.exportToJson();
		        JSONObject jsonObject = new JSONObject(json);
		        JSONObject dataArray = jsonObject.getJSONArray("data").getJSONObject(25);
		        JSONArray imagesArray = dataArray.getJSONArray("images");
		        int len = imagesArray.length();

		        long endTime = System.currentTimeMillis();
		        long elapsedMs = endTime - startTime;

		        assertEquals(5, len);
		        assertTrue(elapsedMs<10000, "The  export cost time shall not greater than 10s");
		    }
		    @Test
		    public void testCELLSGRIDJS404() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp=(new LocalFileCache());

		        // 使用给定的文件名获取测试文件路径
		        String path = getTestfile("autofill.xlsx");
		        try (InputStream fs = new FileInputStream(path)) {
		            // 导入Excel文件
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xlsx"), null);
		        }

		        // 导出为JSON并解析结果
		        String json = gridJsWorkbook.exportToJson();
		        JSONObject cpresult = new JSONObject(json);
		        String filename = cpresult.getString("filename");

		        // 断言文件名是否符合预期
		        assertEquals("book1", filename);

		        // 测试包含非ASCII字符的文件名
		        String nonAsciiFilename = "Информация об объектах, находящихся в муниципальной собственности (ГАСУ) на 01.01.2021.xlsx";
		        json = gridJsWorkbook.exportToJson(nonAsciiFilename);

		        cpresult = new JSONObject(json);
		        filename = cpresult.getString("filename");

		        // 断言非ASCII文件名是否正确设置
		        assertEquals(nonAsciiFilename, filename);
		    }
		    @Test
		    public void testCELLSGRIDJS680() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp = new LocalFileCache();

		        String path = getTestfile("Enron_000008128.xls.xlsx");
		        try (InputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xlsx"), null);
		        }

		        String json = gridJsWorkbook.exportToJson();
		        JSONObject cpresult = new JSONObject(json);
		        String sheetName = ((JSONObject) ((JSONObject) cpresult.getJSONArray("data").get(0))).getString("name");

		        assertEquals("w5000", sheetName);
		    }

		    // Test method for CELLSGRIDJS691
		    @Test
		    public void testCELLSGRIDJS691() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp = new LocalFileCache();

		        String path = getTestfile("Enron_000069227.xls");
		        try (InputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), null);
		        }

		        String json = gridJsWorkbook.exportToJson();
		        JSONObject cpresult = new JSONObject(json);
		        JSONObject sheetData = (JSONObject) cpresult.getJSONArray("data").get(0);
		        String bgColor = sheetData.getString("bgcolor");

		        assertEquals("#339966", bgColor);
		    }

		    // Test method for CELLSGRIDJS690
		    @Test
		    public void testCELLSGRIDJS690() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp = new LocalFileCache();

		        String path = getTestfile("hightlightwords.xlsx");
		        try (InputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xlsx"), null);
		        }

		        String json = gridJsWorkbook.exportToJson();
		        JSONObject cpresult = new JSONObject(json);
		        JSONObject sheetData = (JSONObject) cpresult.getJSONArray("data").get(0);
		        String rowCount = sheetData.getJSONObject("rows").getString("len");

		        assertEquals("21", rowCount);
		    }
		    
		    @Test
		    public void testCELLSAPP1057() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp = new LocalFileCache();

		        String path = getTestfile("lodki.dp.ua export BACK.csv");
		        try (InputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".csv"), null);
		        }

		        String json = gridJsWorkbook.exportToJson();
		        System.out.println(json);
		        // Parse the JSON and get the unique ID for further operations
		        JSONObject cpresult = new JSONObject(json);
		        String uid = cpresult.getString("uniqueid");

		        // Update operation JSON string (this should be constructed as per actual requirements)
		        String p =  "{\"sheetname\":\"sheet1\",\"actrow\":30,\"actcol\":8,\"datas\":[{\"name\":\"sheet1\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"#000000\",\"valign\":\"middle\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false}},{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"custom\":\"+0\"},{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"custom\":\"0+%\"},{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"custom\":\"0%\"},{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"custom\":\"m/yy\"},{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"custom\":\"yy-m\"}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":true,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{},\"rows\":{\"height\":17,\"len\":12166},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}";

		        // Perform the update operation (assuming a method exists in GridJsWorkbook)
		        gridJsWorkbook.mergeExcelFileFromJson(uid, p);

		        // Set the PDF save timeout configuration
		        int MAX_PDF_SAVE_SECONDS = 5;
		        Config.setMaxPdfSaveSeconds(MAX_PDF_SAVE_SECONDS);

		        String filename = "lodki.123.pdf";
		        GridJsWorkbook gwb = new GridJsWorkbook();
		        // Download operation JSON string (this should be constructed as per actual requirements)
		        String pdownload = "{\"sheetname\":\"Sheet1\",\"actrow\":4,\"actcol\":5,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"},{\"textwrap\":false,\"color\":\"#000000\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"},{\"align\":\"left\"}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":false,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{},\"rows\":{\"height\":19,\"len\":13},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"},{\"textwrap\":false,\"color\":\"#000000\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":false,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{},\"rows\":{\"height\":19,\"len\":13},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"#000000\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":false,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{},\"rows\":{\"height\":19,\"len\":13},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}"; // Construct the JSON string for download operation

		        // Perform the download operation (assuming a method exists in GridJsWorkbook)
		        gwb.mergeExcelFileFromJson(uid, pdownload);

		        // Start the stopwatch to measure the save operation duration
		        long startTime = System.currentTimeMillis();
		        gwb.saveToCacheWithFileName(uid, filename, null); // Assuming this method exists and saves the file
		        long endTime = System.currentTimeMillis();
		        long elapsedMs = endTime - startTime;
		        System.out.println("Elapsed: " + elapsedMs);

		        // Assert that the elapsed time is within the expected range
		        assertTrue(elapsedMs >= MAX_PDF_SAVE_SECONDS * 1000, 
		            "The actual time cost for save was less than " + MAX_PDF_SAVE_SECONDS + " seconds");
		        assertTrue(elapsedMs <= (MAX_PDF_SAVE_SECONDS + 3) * 1000, 
		            "The actual time cost for save was more than " + (MAX_PDF_SAVE_SECONDS + 3) + " seconds");
		    }
		    
		    @Test
		    public void testGRIDJS359() throws Exception {
		        Config.setSaveHtmlAsZip(true);
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp = new LocalFileCache();

		        String path = getTestfile("test.htm");
		        try (InputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".htm"), null);
		        }

		        String json = gridJsWorkbook.exportToJson();
		        System.out.println(json);
		        // Do update
		        JSONObject cpresult = new JSONObject(json);
		        String uid = cpresult.getString("uniqueid");
		        String p = "{\"name\":\"Sheet1\",\"ri\":2,\"ci\":1,\"text\":\"cc\",\"op\":\"update\"}";
		        String ret = gridJsWorkbook.updateCell(p, uid);

		        // Do download
		        String filename = "123.htm";
		        GridJsWorkbook gwb = new GridJsWorkbook();
		        String pdownload =  "{\"sheetname\":\"Sheet1\",\"actrow\":4,\"actcol\":5,\"datas\":[{\"name\":\"Sheet1\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"},{\"textwrap\":false,\"color\":\"#000000\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"},{\"align\":\"left\"}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":false,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{},\"rows\":{\"height\":19,\"len\":13},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet2\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"#000000\",\"align\":\"right\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"},{\"textwrap\":false,\"color\":\"#000000\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":false,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{},\"rows\":{\"height\":19,\"len\":13},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]},{\"name\":\"Sheet3\",\"freeze\":\"A1\",\"styles\":[{\"textwrap\":false,\"color\":\"#000000\",\"valign\":\"middle\",\"font\":{\"name\":\"宋体\",\"size\":12,\"bold\":false,\"italic\":false},\"custom\":\"General\"}],\"comments\":[],\"canselectlocked\":true,\"sprotected\":false,\"canselectunlocked\":true,\"showGrid\":false,\"merges\":[],\"settings\":{\"mode\":\"edit\",\"updateMode\":\"server\",\"updateUrl\":\"/GridJs2/UpdateCell\",\"folderName\":\"\",\"fileName\":\"\",\"view\":{},\"showToolbar\":true,\"showPartToolbar\":false,\"showContextmenu\":true,\"row\":{\"len\":100,\"height\":25},\"col\":{\"len\":26,\"width\":100,\"indexWidth\":60,\"minWidth\":60},\"style\":{\"align\":\"left\",\"valign\":\"middle\",\"textwrap\":false,\"strike\":false,\"underline\":false,\"color\":\"#0a0a0a\",\"font\":{\"name\":\"Arial\",\"size\":10,\"bold\":false,\"italic\":false},\"format\":\"normal\"}},\"cols\":{},\"rows\":{\"height\":19,\"len\":13},\"validations\":[],\"autofilter\":{},\"images\":[],\"shapes\":[]}]}"; // Construct the JSON string for download operation
		        gwb.mergeExcelFileFromJson(uid, pdownload);
		        gwb.saveToCacheWithFileName(uid, filename, null);

		        if (filename.endsWith(".html") || filename.endsWith(".htm")) {
		            filename += ".zip";
		        }

		        String downloadPath = GridJsWorkbook.getImageUrl(uid, filename, "/");
		        int ididx = downloadPath.indexOf("id=");
		        String downloadId = downloadPath.substring(ididx + 3);
		        System.out.println(downloadId);

		        InputStream fsReader = (InputStream) gridJsWorkbook.CacheImp.loadStream(downloadId);
		        byte[] bytes = new byte[fsReader.available()];
		        fsReader.read(bytes);
		        String fileOutput = getOutputfile("123htm.zip");
		        if (new File(fileOutput).exists()) {
		            new File(fileOutput).delete();
		        }
		        try (FileOutputStream fwrite = new FileOutputStream(fileOutput)) {
		            fwrite.write(bytes);
		        }

		        String zipOutputDir = getOutputfile("123htm");
		        Util.unZip(fileOutput, zipOutputDir);
		        Thread.sleep(1000);
		        Workbook workbook = new Workbook(getOutputfile("123htm/123.htm"));
		        assertEquals("cc", workbook.getWorksheets().get("Sheet1").getCells().get(2, 1).getStringValue());
		        // Restore config
		        Config.setSaveHtmlAsZip(false);
		    }
		    
		    @Test
		    public void testImageLimit() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp = new LocalFileCache();

		        String path = getTestfile("pictest.xls");
		        try (InputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), null);
		        }

		        Config.setMaxShapeOrImageCount(2);
		        Config.setMaxTotalShapeOrImageCount(31);
		        String json = gridJsWorkbook.exportToJson();
		        System.out.println(json);
		        JSONObject cpresult = new JSONObject(json);
		        int len = ((JSONObject) cpresult.getJSONArray("data").get(0)).getJSONArray("shapes").length();
		        int len2 = ((JSONObject) cpresult.getJSONArray("data").get(1)).getJSONArray("shapes").length();
		        int len3 = ((JSONObject) cpresult.getJSONArray("data").get(2)).getJSONArray("shapes").length();
		        assertEquals(29, len);
		        assertEquals(2, len2);
		        assertEquals(1, len3);

		        len2 = ((JSONObject) cpresult.getJSONArray("data").get(1)).getJSONArray("images").length();
		        len3 = ((JSONObject) cpresult.getJSONArray("data").get(2)).getJSONArray("images").length();
		        assertEquals(2, len2);
		        assertEquals(4, len3);

		        // Do not limit the shape or image count
		        Config.setIslimitShapeOrImage(false);
		        json = gridJsWorkbook.exportToJson();
		        cpresult = new JSONObject(json);
		        System.out.println(json);
		        len = ((JSONObject) cpresult.getJSONArray("data").get(0)).getJSONArray("shapes").length();
		        len2 = ((JSONObject) cpresult.getJSONArray("data").get(1)).getJSONArray("shapes").length();
		        len3 = ((JSONObject) cpresult.getJSONArray("data").get(2)).getJSONArray("shapes").length();
		        assertEquals(29, len);
		        assertEquals(11, len2);
		        assertEquals(2, len3);

		        len2 = ((JSONObject) cpresult.getJSONArray("data").get(1)).getJSONArray("images").length();
		        len3 = ((JSONObject) cpresult.getJSONArray("data").get(2)).getJSONArray("images").length();
		        assertEquals(3, len2);
		        assertEquals(4, len3);

		        // Reset config to default value
		        Config.setIslimitShapeOrImage(true);
		    }

			@Test
			public void testCopyImageToSameSheet() throws Exception {
				GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
				gridJsWorkbook.CacheImp = new LocalFileCache();
				String path = getTestfile("pictest.xls");
				try (InputStream fs = new FileInputStream(path)) {
					gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), "123");
				}
				String json = gridJsWorkbook.exportToJson();
				JSONObject cpresult = new JSONObject(json);
				String uid = cpresult.getString("uniqueid");

				String p = "{\"name\":\"Sheet2\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":3,\"srcid\":146,\"isshape\":false}";
				String p2 = "{\"name\":\"Sheet2\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":1,\"srcid\":100,\"isshape\":false}";
				String expect_w = "206";
				String expect_h = "72";
				String expect_err_msg = "{\"Error\":\"wrong picture id \"}";

				Workbook wb1 = new Workbook(path);
				int count = wb1.getWorksheets().get("Sheet2").getPictures().getCount();
				testAddPictureOrShape(gridJsWorkbook, uid, p, p2, expect_w, expect_h, expect_err_msg, count, "Sheet2");

			}
			
			@Test
		    public void testCopyImageToOtherSheet() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp = new LocalFileCache();
		        String path = getTestfile("pictest.xls");

		        try (InputStream fs = new FileInputStream(path)) {
		            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), "123");
		        }
		        String json = gridJsWorkbook.exportToJson();
		        System.out.println(json);
		        JSONObject cpresult = new JSONObject(json);
		        String uid = cpresult.getString("uniqueid");

		        String p = "{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":3,\"srcid\":146,\"isshape\":false}";
		        String p2 = "{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":1,\"srcid\":100,\"isshape\":false}";
		        String expect_w = "206";
		        String expect_h = "72";
		        String expect_err_msg = "{\"Error\":\"wrong picture id \"}";

		        Workbook wb1 = new Workbook(path); // No try-with-resources
		        int count = wb1.getWorksheets().get("Sheet3").getPictures().getCount();
		        testAddPictureOrShape(gridJsWorkbook, uid, p, p2, expect_w, expect_h, expect_err_msg, count, "Sheet3");
		        
		    }
			@Test
		    public void testCopyImageWithNoStreamCacheImp() throws Exception {
		        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
		        gridJsWorkbook.CacheImp =null;
		        String path = getTestfile("pictest.xls");
		        gridJsWorkbook.importExcelFile(path);

		        String json = gridJsWorkbook.exportToJson();
		        System.out.println(json);
		        JSONObject cpresult = new JSONObject(json);
		        String uid = cpresult.getString("uniqueid");

		        String p = "{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":3,\"srcid\":146,\"isshape\":false}";
		        String p2 = "{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":1,\"srcid\":100,\"isshape\":false}";
		        String expect_w = "206";
		        String expect_h = "72";
		        String expect_err_msg = "{\"Error\":\"wrong picture id \"}";

		        Workbook wb1 = new Workbook(path); // No try-with-resources
		        int count = wb1.getWorksheets().get("Sheet3").getPictures().getCount();
		        testAddPictureOrShape(gridJsWorkbook, uid, p, p2, expect_w, expect_h, expect_err_msg, count, "Sheet3");
		        
		    }
			
			    @Test
			    public void testCELLSGRIDJS392ReloadByUid() throws Exception {
			        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
			        gridJsWorkbook.CacheImp = new LocalFileCache();
			        ByteArrayOutputStream ms = new ByteArrayOutputStream();
			        Workbook wb = new Workbook();
			        wb.save(ms, SaveFormat.XLSX);
			        gridJsWorkbook.importExcelFile(new java.io.ByteArrayInputStream(ms.toByteArray()), GridLoadFormat.XLSX);
			            
			       
			        String json = gridJsWorkbook.exportToJson();
			        JSONObject cpresult = new JSONObject(json);
			        String uid = cpresult.getString("uniqueid");
			        String p = "{\"name\":\"Sheet1\",\"ri\":4,\"ci\":11,\"text\":\"=sin(1)\",\"op\":\"update\"}";
			        gridJsWorkbook.updateCell(p, uid);
			        StringBuilder ret = gridJsWorkbook.getJsonByUid(uid, "hello1.xlsx");
			        System.out.println(ret.toString()); // Use the indented print method if needed

			        String path = getTestfile("simplejava.json");
			        String expected = new String(java.nio.file.Files.readAllBytes(java.nio.file.Paths.get(path)));
			        JSONObject cpresult2 = new JSONObject(ret.toString());
			        System.out.println(cpresult2.getJSONArray("data").getString(0));

			        assertEquals(expected, cpresult2.getJSONArray("data").getString(0));

			        // Test for a wrong uid
			        StringBuilder ret3 = null;
			        try {
			            ret3 = gridJsWorkbook.getJsonByUid(uid + "123", "");
			        } catch (IOException e) {
			            // If an exception is expected, handle it here
			        }
			        assertNull(ret3);
			    }
			    
			    @Test
				public void testCELLSGRIDJS433Sort() throws Exception {
					GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
					gridJsWorkbook.CacheImp = new LocalFileCache();

					String path = getTestfile("travel-cccbudget.xlsx");
					InputStream fs = new FileInputStream(path);
					Workbook workbook = new Workbook(fs);
					ByteArrayOutputStream ms = new ByteArrayOutputStream();
					workbook.save(ms, SaveFormat.XLSX);
					gridJsWorkbook.importExcelFile(new java.io.ByteArrayInputStream(ms.toByteArray()),
							GridLoadFormat.XLSX);

					String json = gridJsWorkbook.exportToJson();
					System.out.println(json);
					JSONObject cpresult = new JSONObject(json);
					String uid = cpresult.getString("uniqueid");

					// Sort with header
					String p = "{\"name\":\"Travel Budget\",\"src\":{\"sri\":13,\"sci\":1,\"eri\":20,\"eci\":6,\"w\":0,\"h\":0},\"ci\":[5],\"order\":[\"asc\"],\"isheader\":true,\"op\":\"sort\"}";
					String ret = gridJsWorkbook.updateCell(p, uid);
					String expected = "{\"op\":\"sort\",\"status\":\"ok\",\"r\":{\"sri\":14,\"eri\":20,\"sci\":1,\"eci\":6},\"rowids\":[17,18,16,20,19,15,14]}";
					assertEquals(expected, ret);

					// Sort without header
					p = "{\"name\":\"Travel Budget\",\"src\":{\"sri\":13,\"sci\":3,\"eri\":19,\"eci\":5,\"w\":0,\"h\":0},\"ci\":[2],\"order\":[\"desc\"],\"isheader\":false,\"op\":\"sort\"}";
					expected = "{\"op\":\"sort\",\"status\":\"ok\",\"r\":{\"sri\":13,\"eri\":19,\"sci\":3,\"eci\":5},\"rowids\":[13,18,16,19,17,14,15]}";
					ret = gridJsWorkbook.updateCell(p, uid);
					assertEquals(expected, ret);
				}
			    
				@Test
				public void testCopyShape() throws Exception {
					GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
					gridJsWorkbook.CacheImp = new LocalFileCache();

					String path = getTestfile("pictest.xls");
					try (InputStream fs = new FileInputStream(path)) {
						gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), "123");
					}
					String json = gridJsWorkbook.exportToJson();
					System.out.println(json);
					JSONObject cpresult = new JSONObject(json);
					String uid = cpresult.getString("uniqueid");

					String p = "{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":3,\"srcid\":135,\"isshape\":true}";
					String p2 = "{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":1,\"srcid\":100,\"isshape\":true}";
					String expect_w = "138";
					String expect_h = "46";
					String expect_err_msg = "{\"Error\":\"wrong shape id \"}";

					Workbook wb1 = new Workbook(path);
					int count = wb1.getWorksheets().get("Sheet3").getShapes().getCount();
					testAddPictureOrShape(gridJsWorkbook, uid, p, p2, expect_w, expect_h, expect_err_msg, count,
							"Sheet3");

				}
				
				@Test
			    public void testCopySameShapeMultipleTimes() throws Exception {
			        GridJsWorkbook gridJsWorkbook = new GridJsWorkbook();
			        gridJsWorkbook.CacheImp = new LocalFileCache();

			        String path = getTestfile("pictest.xls");
			        try (InputStream fs = new FileInputStream(path)) {
			            gridJsWorkbook.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat(".xls"), "123");
			        }
			        String json = gridJsWorkbook.exportToJson();
			        System.out.println(json);
			        JSONObject cpresult = new JSONObject(json);
			        String uid = cpresult.getString("uniqueid");
			        int ccc = 0;

			        for (int ri = 2; ri <= 12; ri++) {
			            for (int ci = 3; ci <= 10; ci++) {
			                String p = String.format("{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":%d,\"ci\":%d,\"srcid\":1,\"isshape\":true}", ri, ci);
			                String p2 = "{\"name\":\"Sheet3\",\"srcname\":\"Sheet2\",\"ri\":11,\"ci\":1,\"srcid\":100,\"isshape\":true}";
			                String expect_w = "141";
			                String expect_h = "49";
			                String expect_err_msg = "{\"Error\":\"wrong shape id \"}";

			                // Copy the shape multiple times
			                try {
			                    String picret = gridJsWorkbook.copyImageOrShape(uid, p);
			                    picret = gridJsWorkbook.copyImageOrShape(uid, p);
			                    picret = gridJsWorkbook.copyImageOrShape(uid, p);
			                    ccc += 3;
			                } catch (Exception e) {
			                    System.out.println(e.getMessage());
			                }
			            }
			            System.out.println("execute copy :" + ccc);
			        }
			    }
		    
		    //TODO still left 6 cases
				
		    private static void testAddPictureOrShape(GridJsWorkbook gw, String uid, String p, String p2, String expect_w, String expect_h, String expect_err_msg, int expectid, String targetsheet) throws Exception {
		        String picret = gw.copyImageOrShape(uid, p);

		        JSONObject paramjson = new JSONObject(p);
		        boolean isshape = paramjson.getBoolean("isshape");

		        System.out.println("copy image :" + picret);
		        JSONObject picretjson = new JSONObject(picret);
		        int picid = picretjson.getInt("id");
		        assertEquals(expectid, picid);

		        String picurl = picretjson.getString("url");
		        if (picurl == null) {
		            // Handle the case when picurl is null
		        } else {
		            int ididx = picurl.indexOf("id=");
		            String downloadid = picurl.substring(ididx + 3);
		            System.out.println(downloadid);
		            // Additional code to test the image stream might go here
		        }

		        ByteArrayOutputStream ms = new ByteArrayOutputStream();
		        gw.saveToExcelFile(ms);
		        ms.flush();
		        ByteArrayInputStream inputStream = new ByteArrayInputStream(ms.toByteArray());
		        Workbook wb = new Workbook(inputStream);
		        Worksheet sheet = wb.getWorksheets().get(targetsheet);
		        Shape sp = isshape ? sheet.getShapes().get(picid) : sheet.getPictures().get(picid);

		        int w = sp.getWidth();
		        int h = sp.getHeight();

		        assertEquals(expect_w, String.valueOf(w));
		        assertEquals(expect_h, String.valueOf(h));
		        assertEquals(expect_w, String.valueOf(picretjson.getInt("width")));
		        assertEquals(expect_h, String.valueOf(picretjson.getInt("height")));
		        assertEquals(11, sp.getUpperLeftRow());
		        assertEquals(3, sp.getUpperLeftColumn());

		        picret = gw.copyImageOrShape(uid, p2);
		        assertEquals(expect_err_msg, picret);
		    }
		    
		    public static InputStream getStreamFromUrl(String url) throws Exception {
		        HttpsURLConnection connection = (HttpsURLConnection) new URL(url).openConnection();
		        connection.setRequestMethod("GET");
		        connection.connect();

		        int responseCode = connection.getResponseCode();
		        if (responseCode == HttpsURLConnection.HTTP_OK) {
		            return connection.getInputStream();
		        } else {
		            throw new RuntimeException("Failed to download the image: HTTP error code " + responseCode);
		        }
		    }
		    private void insertImage(String url, String[] uidOut) throws Exception {
		    	GridJsWorkbook gw = new GridJsWorkbook();
		        gw.CacheImp=(new LocalFileCache()); // Use stream cache

		        String path = getTestfile("pictest.xls");
		        try (FileInputStream fs = new FileInputStream(path)) {
		            gw.importExcelFile(fs, GridJsWorkbook.getGridLoadFormat("pictest.xls"), "123");
		        }
		        String json = gw.exportToJson();
		        System.out.println(json);
		        JSONObject cpresult = new JSONObject(json);
		        String uid = cpresult.getString("uniqueid");
		        uidOut[0] = uid;
		        String p = "{\"name\":\"Sheet2\",\"ri\":1,\"ci\":1}";
		        String picRet = null;
		        String expectW = null;
		        String expectH = null;

		        InputStream imageStream = null;
		        if (url == null) { // Use local image file
		            File imageFile = new File(getTestfile("snap1.png"));
		            try (FileInputStream s = new FileInputStream(imageFile)) {
		                imageStream = s;
		                picRet = gw.insertImage(uid, p, imageStream, url);
		                expectW = "320";
		                expectH = "204";
		            }
		        } else { // Use URL to get image
		            imageStream = getStreamFromUrl(url);
		            picRet = gw.insertImage(uid, p, imageStream, url);
		            expectW = "202";
		            expectH = "66";
		        }
		        System.out.println("insert first image : " + picRet);
		        JSONObject picRetJson = new JSONObject(picRet);

		        int picId = picRetJson.getInt("id");
		        assertEquals(3, picId);

		        ByteArrayOutputStream baos = new ByteArrayOutputStream();
		        gw.saveToExcelFile(baos);

		        baos.close();

		        byte[] byteArray = baos.toByteArray();
		        try (ByteArrayInputStream bais = new ByteArrayInputStream(byteArray)) {
		            Workbook wb = new Workbook(bais);

		            Worksheet sheet = wb.getWorksheets().get("Sheet2"); // Assuming "Sheet2" is the name of the sheet
		            Shape sp = (Shape) sheet.getPictures().get(picId); // Shapes are typically stored in a list of sheet objects, and the ID might need to be adjusted

		            int w = sp.getWidth();
		            int h = sp.getHeight();

		            assertEquals(expectW, Integer.toString(w));
		            assertEquals(expectH, Integer.toString(h));
		            assertEquals(expectW, picRetJson.getString("width"));
		            assertEquals(expectH, picRetJson.getString("height"));

		            // Mock another post from controller action
		            // Add second image with a new GridJsWorkbook instance
		            gw = new GridJsWorkbook();
		            gw.CacheImp=(new LocalFileCache()); // Use stream cache
		            p = "{\"name\":\"Sheet2\",\"ri\":3,\"ci\":3}";

		            if (url == null) { // Use local image file
		                File imageFile = new File(getTestfile("snap1.png"));
		                try (FileInputStream s = new FileInputStream(imageFile)) {
		                    picRet = gw.insertImage(uid, p, s, url);
		                    System.out.println("insert second image : " + picRet);
		                }
		            } else { // Use URL to get image
		                imageStream = getStreamFromUrl(url);
		                picRet = gw.insertImage(uid, p, imageStream, url);
		            }

		            picRetJson = new JSONObject(picRet);

		            baos = new ByteArrayOutputStream();
		            gw.saveToExcelFile(baos);

		            baos.close();

		            byteArray = baos.toByteArray();
		            try (ByteArrayInputStream bais2 = new ByteArrayInputStream(byteArray)) {
		                Workbook wb2 = new Workbook(bais2);
		                picId = picRetJson.getInt("id");
		                assertNotNull(picRetJson.get("url"));
		                assertEquals(4, picId);
		                if (url != null) {
		                    assertEquals(url, picRetJson.getString("url"));
		                }
		                
		                sheet = wb2.getWorksheets().get("Sheet2");
		                sp = (Shape) sheet.getPictures().get(picId);
		                w = sp.getWidth();
		                h = sp.getHeight();
		                assertEquals(expectW, Integer.toString(w));
		                assertEquals(expectH, Integer.toString(h));

		                // The first picture shall still be there
		                sp = (Shape) sheet.getPictures().get(3);  
		                w = sp.getWidth();
		                h = sp.getHeight();
		                assertEquals(expectW, Integer.toString(w));
		                assertEquals(expectH, Integer.toString(h));
		            }
		        }
		    }
		    
		    private static void testImageZipFiles(String uid, int expectedFileCount) throws IOException {
		    	GridJsWorkbook gw = new GridJsWorkbook();
		        String zipOutDir = getOutputfile("s0_batch");
		        String outputFile = getOutputfile("s0_batch.zip");

		        try {
		            Files.deleteIfExists(Paths.get(outputFile));
		            Util.clearFolder((zipOutDir));
		        } catch (IOException e) {
		            System.out.println("Delete failed: " + e.getMessage());
		        }

		        String filename = "s0_batch.zip";
		        String downloadPath = GridJsWorkbook.getImageUrl(uid, filename, "/");
		        int idIdx = downloadPath.indexOf("id=");
		        String downloadId = downloadPath.substring(idIdx + 3);
		        System.out.println(downloadId);

		        try (InputStream fsReader = gw.CacheImp.loadStream(downloadId)) {
		            
		            try (FileOutputStream fwrite = new FileOutputStream(outputFile)) {
		                Util.copyStream(fsReader, fwrite);
		            }
		        }

		        Util.unZip(outputFile, zipOutDir);

		        Path dirPath = Paths.get(zipOutDir);
		        File[] files = dirPath.toFile().listFiles();
		         assertEquals(expectedFileCount, files.length);
		    }

		    private static void addPictureToWb(Random rd, Workbook wb, String imagePath) throws Exception {
		        try (FileInputStream s = new FileInputStream(imagePath)) {
		            for (int i = 0; i <= 400; i += 40) {
		                for (int j = 0; j <= 100; j += 5) {
		                    int picId = wb.getWorksheets().get(0).getPictures().add( i, j,s);
		                    wb.getWorksheets().get(0).getCells().get(i, j).putValue(i + " - " + j);
		                    if (picId % 10 == 0) {
		                        wb.getWorksheets().get(0).getPictures().get(picId).setRotationAngle(rd.nextInt(180));
		                    }
		                }
		            }
		        }
		    }
}
