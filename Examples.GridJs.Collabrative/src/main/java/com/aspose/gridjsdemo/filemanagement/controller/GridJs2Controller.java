package com.aspose.gridjsdemo.filemanagement.controller;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.zip.GZIPOutputStream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
//import jakarta.servlet.http.HttpServletRequest;
//import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.aspose.gridjs.GridJsControllerBase;
import com.aspose.gridjs.IGridJsService;



@RestController
@RequestMapping({"/GridJs2"})
public class GridJs2Controller extends GridJsControllerBase{
     @Value("${fileconfig.ListDir}")
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
     
     
     @Value("${fileconfig.UploadPath}")
     private String uploadDir;
     
     @GetMapping("/DetailStreamJsonWithUidFromUpload")
     public void detailStreamJsonWithUidFromUpload(@RequestParam String filename, @RequestParam String uid,HttpServletResponse response) {
        
            
         	Path filePath = Paths.get(uploadDir, filename);
           

             response.setContentType("application/json");
             response.setHeader("Content-Encoding", "gzip");
             try (GZIPOutputStream gzipOutputStream = new GZIPOutputStream(response.getOutputStream())) {
             	   _gridJsService.detailStreamJsonWithUid(gzipOutputStream, filePath.toString(), uid);
 				 
             }  catch (Exception e) {
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
         return super.updateCell(request);
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