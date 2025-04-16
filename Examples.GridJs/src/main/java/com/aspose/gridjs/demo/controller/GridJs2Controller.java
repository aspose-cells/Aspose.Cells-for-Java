package com.aspose.gridjs.demo.controller;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.zip.GZIPOutputStream;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import com.aspose.cells.LoadFormat;
import com.aspose.cells.Workbook;
import com.aspose.gridjs.Config;
import com.aspose.gridjs.GridCellException;
import com.aspose.gridjs.GridInterruptMonitor;
import com.aspose.gridjs.GridJsControllerBase;
import com.aspose.gridjs.GridJsWorkbook;
import com.aspose.gridjs.IGridJsService;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;

@RestController
@RequestMapping({"/GridJs2"})
public class GridJs2Controller extends GridJsControllerBase{
     @Value("${testconfig.ListDir}")
    private String listDir;//="D:\\codebase\\customerissue\\wb\\tempfromdownload\\";

    
  //   IGridJsService
     @Autowired
     public GridJs2Controller(IGridJsService gridJsService) {
         super(gridJsService);  
     }
     
 
    @GetMapping("/DetailFileJsonWithUid")
    public ResponseEntity<String> detailFileJsonWithUid(@RequestParam String filename, @RequestParam String uid) {
        try {
           
        	Path filePath = Paths.get(listDir, filename);

            // Check if already in cache
            StringBuilder sb =   _gridJsService.detailFileJsonWithUid(  filePath.toString(), uid);

            // Return the content as plain text with UTF-8 encoding
            return ResponseEntity.ok()
                    .header("Content-Type", "text/plain; charset=UTF-8")
                    .body(sb.toString());
        } catch (Exception e) {
            return ResponseEntity.status(500).body("Error processing the file: " + e.getMessage());
        }
    }
    
    @GetMapping("/DetailStreamJsonWithUid")
    public void detailStreamJsonWithUid(@RequestParam String filename, @RequestParam String uid,HttpServletResponse response) {
       
           
        	Path filePath = Paths.get(listDir, filename);
            

            response.setContentType("application/json");
            response.setHeader("Content-Encoding", "gzip");
            try (GZIPOutputStream gzipOutputStream = new GZIPOutputStream(response.getOutputStream())) {
            	  _gridJsService.detailStreamJsonWithUid(gzipOutputStream, filePath.toString(), uid);
            } catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
    }
    
    
    @Value("${testconfig.UploadPath}")
    private String uploadDir;
    
    @GetMapping("/DetailStreamJsonWithUidFromUpload")
    public void detailStreamJsonWithUidFromUpload(@RequestParam String filename, @RequestParam String uid,HttpServletResponse response) {
       
           
        	Path filePath = Paths.get(uploadDir, filename);
          

            response.setContentType("application/json");
            response.setHeader("Content-Encoding", "gzip");
            try (GZIPOutputStream gzipOutputStream = new GZIPOutputStream(response.getOutputStream())) {
            	   _gridJsService.detailStreamJsonWithUid(gzipOutputStream, filePath.toString(), uid);
				 
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
    }
    
    @PostMapping("/LazyLoadingStreamJson")
    public void lazyLoadingStreamJson(
            @RequestParam(value = "name", required = false) String sheetName,
            @RequestParam(value = "uid", required = false) String uid,
            HttpServletResponse response) throws IOException {


         super.lazyLoadingStreamJson(sheetName,uid,response);

    }
    
    @PostMapping("/UpdateCell")
    public ResponseEntity<String> updateCell(HttpServletRequest request) {
		try {
        return super.updateCell(request);
    	}catch(Exception e)
    	{
    		e.printStackTrace();
    		return null;
		}
	}
	
	@PostMapping("/AddImage")
    public ResponseEntity<String> addImage(
            @RequestParam(value = "image", required = false) MultipartFile file,
			@RequestParam("uid") String uid, 
			@RequestParam("p") String p,
			@RequestParam(value = "control", required = false) String isControl) {
						
    	 return super.addImage(file,uid,p,isControl);
	}

    @PostMapping("/CopyImage")
    public ResponseEntity<String> copyImage(HttpServletRequest request) {
    	 return super.copyImage(request);
    }


    @PostMapping("/AddImageByURL")
    public ResponseEntity<String> addImageByUrl(HttpServletRequest request) {
  		 return super.addImageByUrl(request);
    }
    
    @GetMapping("/Image")
    public ResponseEntity<InputStreamResource> getImage(HttpServletRequest request) {
    	 return super.getImage(request);
    }
    
    
    @GetMapping("/ImageUrl")
    public ResponseEntity<String> getImageUrl(@RequestParam String id, @RequestParam String uid) {
      return super.getImageUrl(id, uid);
        }

    @GetMapping("/Ole")
    public ResponseEntity<?> getOle(HttpServletRequest request) {
       return super.getOle(request);
    }
    
    @GetMapping("/GetFile")
    public ResponseEntity<?> getFile(@RequestParam("id") String id) {
        return super.getFile(id);
    }
    
    @PostMapping("/Download")
	public ResponseEntity<String> download(HttpServletRequest request) {
		 return super.download(request);
	}





 
}
