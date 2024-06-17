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

import com.aspose.cells.Workbook;
import com.aspose.gridjs.Config;
import com.aspose.gridjs.GridCellException;
import com.aspose.gridjs.GridInterruptMonitor;
import com.aspose.gridjs.GridJsWorkbook;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.Part;

@RestController
@RequestMapping({"/GridJs2"})
public class GridJs2Controller {
     @Value("${testconfig.ListDir}")
    private String listDir;//="D:\\codebase\\customerissue\\wb\\tempfromdownload\\";


    @GetMapping("/DetailFileJsonWithUid")
    public ResponseEntity<String> detailFileJsonWithUid(@RequestParam String filename, @RequestParam String uid) {
        try {
           
        	Path filePath = Paths.get(listDir, filename);
            GridJsWorkbook wbj = new GridJsWorkbook();

            // Check if already in cache
            StringBuilder sb = wbj.getJsonByUid(uid, filename);
            if (sb == null) {
                Workbook wb = new Workbook(filePath.toString());
                wbj.importExcelFile(uid, wb);
                sb = wbj.exportToJsonStringBuilder(filename);
            }

            // Return the content as plain text with UTF-8 encoding
            return ResponseEntity.ok()
                    .header("Content-Type", "text/plain; charset=UTF-8")
                    .body(sb.toString());
        } catch (Exception e) {
            return ResponseEntity.status(500).body("Error processing the file: " + e.getMessage());
        }
    }
    
    @PostMapping("/UpdateCell")
    public ResponseEntity<String> updateCell(HttpServletRequest request) {
        String p = request.getParameter("p");
        String uid = request.getParameter("uid");
        GridJsWorkbook gwb = new GridJsWorkbook();
        String ret;
		try {
			ret = gwb.updateCell(p, uid);
			return new ResponseEntity<>(ret, HttpStatus.OK);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.OK);
		}

    }
    
	@PostMapping("/AddImage_FAIL")
	public ResponseEntity<String> addImageFail(HttpServletRequest request) {
		String uid = request.getParameter("uid");
		String p = request.getParameter("p");
		String isControl = request.getParameter("control");
		String ret = null;
		GridJsWorkbook gwb = new GridJsWorkbook();
		try {
			if (isControl == null) {
				Part filePart = request.getPart("image"); // Retrieves <input type="file" name="file">
				if (filePart == null || filePart.getSize() == 0) {
					// for add shape,need to set p.type as one of AutoShapeType
					ret = gwb.insertImage(uid, p, null, null);

				} else {
					try (InputStream inputStream = filePart.getInputStream()) {
						ret = gwb.insertImage(uid, p, inputStream, null);
					} catch (IOException e) {
						return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.BAD_REQUEST);
					}
				}
			} else {
				ret = gwb.insertImage(uid, p, null, null);
			}
			return new ResponseEntity<>(ret, HttpStatus.OK);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.OK);
		}
	}
	
	@PostMapping("/AddImage")
	public ResponseEntity<String> handleFileUpload(@RequestParam(value = "image", required = false) MultipartFile file,
			@RequestParam("uid") String uid, 
			@RequestParam("p") String p,
			@RequestParam(value = "control", required = false) String isControl) {
		String ret = null;
		GridJsWorkbook gwb = new GridJsWorkbook();
		try {
			if (isControl == null) {
				if (file != null && !file.isEmpty()) {
					try (InputStream inputStream = file.getInputStream()) {
//						Files.copy(inputStream,  Paths.get("D:\\tmpdel\\tmp\\gridjsjava\\upload\\123.png"));
						ret = gwb.insertImage(uid, p, inputStream, null);
						
					} catch (IOException e) {
						return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.BAD_REQUEST);
					}
					

				} else {
					// for add shape,need to set p.type as one of AutoShapeType
					ret = gwb.insertImage(uid, p, null, null);

				}
			} else {
				ret = gwb.insertImage(uid, p, null, null);
			}

			return new ResponseEntity<>(ret, HttpStatus.OK);
		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.internalServerError().body("Error occurred while uploading file");
		}
	}

    @PostMapping("/CopyImage")
    public ResponseEntity<String> copyImage(HttpServletRequest request) {
        String uid = request.getParameter("uid");
        String p = request.getParameter("p");
        GridJsWorkbook gwb = new GridJsWorkbook();
        try {
        String ret = gwb.copyImageOrShape(uid, p);

        return new ResponseEntity<>(ret, HttpStatus.OK);
        } catch (Exception e) {
			// TODO Auto-generated catch block
			return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.OK);
		}
    }

    private static InputStream getStreamFromUrl(String url) throws IOException {
        URL imageUrl = new URL(url);
        URLConnection connection = imageUrl.openConnection();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int bytesRead;
        try (InputStream in = connection.getInputStream()) {
            while ((bytesRead = in.read(buffer)) != -1) {
                baos.write(buffer, 0, bytesRead);
            }
        }
        return new ByteArrayInputStream(baos.toByteArray());
    }

    @PostMapping("/AddImageByURL")
    public ResponseEntity<String> addImageByUrl(HttpServletRequest request) {
        String uid = request.getParameter("uid");
        String p = request.getParameter("p");
        String imageUrl = request.getParameter("imageurl");
        String ret = null;
        GridJsWorkbook gwb = new GridJsWorkbook();
        try {
        if (imageUrl != null) {
            try (InputStream stream = getStreamFromUrl(imageUrl)) {
                ret = gwb.insertImage(uid, p, stream, imageUrl);
            } catch (IOException e) {
                return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.BAD_REQUEST);
            }
        } else {
            return new ResponseEntity<>(gwb.errorJson("image url is null"), HttpStatus.BAD_REQUEST);
        }

        return new ResponseEntity<>(ret, HttpStatus.OK);
        } catch (Exception e) {
			// TODO Auto-generated catch block
			return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.OK);
		}
    }
    
    @GetMapping("/Image")
    public ResponseEntity<InputStreamResource> getImage(HttpServletRequest request) {
        String fileId = request.getParameter("id");
        String uid = request.getParameter("uid");
        try {
        InputStream imageStream = GridJsWorkbook.getImageStream(uid, fileId);
        if (imageStream == null) {
            return ResponseEntity.notFound().build();
        }

        return ResponseEntity.ok()
                .contentType(MediaType.IMAGE_PNG)
                .body(new InputStreamResource(imageStream));
        } catch (Exception e) {
			 e.printStackTrace();
        	 return ResponseEntity.notFound().build();
		}
    }
    /*
    @GetMapping("/ImageTest")
    public ResponseEntity<InputStreamResource> getImageTest(HttpServletRequest request) {
        try {
            // Load the image file from the classpath (resources folder)
        	  String fileId = request.getParameter("id");
              String uid = request.getParameter("uid");
              
              InputStream imageStream = GridJsWorkbook.getImageStream(uid, fileId);
              if (imageStream == null) {
                  return ResponseEntity.notFound().build();
              }

            // Set the headers for the response
            HttpHeaders headers = new HttpHeaders();
            headers.add("Content-Type", "image/jpeg");

            // Return the InputStreamResource with the headers
            return ResponseEntity
                    .ok()
                    .headers(headers)
                    .body(new InputStreamResource(imageStream));
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
    
    @GetMapping("/PImage")
    public ResponseEntity<InputStreamResource> getPImage(HttpServletRequest request) {
        String fileId = request.getParameter("id");
        String uid = request.getParameter("uid");
        try {
        InputStream imageStream = GridJsWorkbook.getImageStream(uid, fileId);
        if (imageStream == null) {
            return ResponseEntity.notFound().build();
        }
        InputStreamResource resource = new InputStreamResource(imageStream);
        MediaType mediaType = MediaType.IMAGE_PNG; 
        return ResponseEntity.ok()
                .contentType(mediaType)
                .header(HttpHeaders.CONTENT_DISPOSITION, "inline; filename=\"" + fileId + "\"")
                .body(resource);
        } catch (Exception e) {
			 e.printStackTrace();
        	 return ResponseEntity.notFound().build();
		}
    }
    
    @GetMapping("/pimages/{imageName}")
    public ResponseEntity<InputStreamResource> getImage(@PathVariable("imageName") String imageName) {
        try {
            // Read the image file
            FileInputStream fis = new FileInputStream("C:\\Users\\peter\\Pictures\\test\\" + imageName);
            InputStreamResource resource = new InputStreamResource(fis);

            // Set the content type based on the image type (JPEG, PNG, etc.)
            MediaType mediaType = MediaType.IMAGE_JPEG; // Change this based on your image type

            // Create the response entity with the appropriate headers
            return ResponseEntity.ok()
                    .contentType(mediaType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "inline; filename=\"" + imageName + "\"")
                    .body(resource);
        } catch (IOException e) {
            // Handle the exception, e.g., log it
            e.printStackTrace();
            return ResponseEntity.notFound().build();
        }
    }
    */
    
    @GetMapping("/Ole")
    public ResponseEntity<?> getOle(HttpServletRequest request) {
        int oleId = Integer.parseInt(request.getParameter("id"));
        String uid = request.getParameter("uid");
        String sheet = request.getParameter("sheet");

        GridJsWorkbook gwb = new GridJsWorkbook();
        String[] filename = new String[1];
        try {
        byte[] fileByte = gwb.getOle(uid, sheet, oleId, filename);
      
        if (filename != null) {
            return ResponseEntity.ok()
                    .contentType( (getMimeType(filename[0])))
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename[0] + "\"")
                    .body(fileByte);
        } else {
            return ResponseEntity.notFound().build();
        }
        } catch (Exception e) {
			// TODO Auto-generated catch block
			return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.OK);
		}
    }
    
    
    @GetMapping("/ImageUrl")
    public ResponseEntity<String> getImageUrl(@RequestParam String id, @RequestParam String uid) {
        if (GridJsWorkbook.CacheImp != null) {
            return ResponseEntity.ok(GridJsWorkbook.getImageUrl(uid, id, "."));
        } else {
            String file = uid + "." + id;
            return ResponseEntity.ok("/GridJs2/GetZipFile?f=" + file);
        }
    }

    @GetMapping("/GetZipFile")
    public ResponseEntity<StreamingResponseBody> getZipFile(@RequestParam String f) {
        Path filePath = Paths.get(Config.getFileCacheDirectory(), f);

        if (!Files.exists(filePath)) {
            return ResponseEntity.notFound().build();
        }

        StreamingResponseBody responseBody = outputStream -> {
            try (InputStream inputStream = Files.newInputStream(filePath)) {
                byte[] buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    outputStream.write(buffer, 0, bytesRead);
                }
            }
        };

        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + f + "\"")
                .body(responseBody);
    }
    
    public static MediaType getMimeType(String fileName) {
        Path filePath = Paths.get(fileName);
        try {
            // Probe the content type of the file
            String mimeType = Files.probeContentType(filePath);
            
            // If the MIME type couldn't be determined, default to binary
            if (mimeType == null) {
                mimeType = "application/octet-stream";
            }

            // Parse and return the MediaType
            return MediaType.parseMediaType(mimeType);

        } catch (IOException e) {
            e.printStackTrace();
            // Return a default MediaType in case of an error
            return MediaType.APPLICATION_OCTET_STREAM;
        }
    }
    
    @GetMapping("/GetFile")
    public ResponseEntity<byte[]> getFile(@RequestParam("id") String fileid,@RequestParam("filename") String filename) throws IOException {
       
        MediaType mimeType = getMimeType(filename);
       
//        try {
//            TimeUnit.SECONDS.sleep(3); // simulate network lag
//        } catch (InterruptedException e) {
//            Thread.currentThread().interrupt();
//        }
        Path filePath = Paths.get(Config.getFileCacheDirectory(), fileid.replace('/', '.')+"."+filename);
        byte[] bytes = Files.readAllBytes(filePath);
        return ResponseEntity.ok()
                .contentType(mimeType)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .body(bytes);
    }
    
    
    //this is for the way that implment GridCacheForStream
    @GetMapping("/GetFileUseCacheStream")
    public ResponseEntity<InputStreamResource>  getFileUseCacheStream(@RequestParam("id") String fileid) throws IOException {
 
    	MediaType mimeType = getMimeType(fileid);
    	String name =   fileid.replace('/', '.');
 
    	
    	 InputStream inputStream = GridJsWorkbook.CacheImp.loadStream(fileid);
         if (inputStream == null) {
             // Handle the case where the file is not found
             return ResponseEntity.notFound().build();
         }

         HttpHeaders headers = new HttpHeaders();
         headers.setContentType(mimeType);
         headers.setContentDispositionFormData("attachment", name);

         return ResponseEntity.ok()
                 .headers(headers)
                 .body(new InputStreamResource(inputStream));
  
    }
    
    
    
    private static void InterruptMonitor(GridInterruptMonitor monitor,int milliseconds)
    {
    	 
        try
        {
            Thread.sleep(milliseconds);
           
            (monitor).interrupt();
        }
        catch (Exception e)
        {
            System.out.println("Succeeded for load in give time.");
        }
    }
    
    @PostMapping("/Download")
	public ResponseEntity<String> download(HttpServletRequest request) {
		String p = request.getParameter("p");
		String uid = request.getParameter("uid");
		String filename = request.getParameter("file");

		GridJsWorkbook gwb = new GridJsWorkbook();
		try {
			gwb.mergeExcelFileFromJson(uid, p);

			GridInterruptMonitor m = new GridInterruptMonitor();
			gwb.setInterruptMonitorForSave(m);

			ExecutorService executor = Executors.newSingleThreadExecutor();
			executor.submit(() -> InterruptMonitor(m, 30 * 1000));

			try {
				gwb.saveToCacheWithFileName(uid, filename, null);
			} catch (Exception ex) {
				if (ex instanceof GridCellException) {
					return ResponseEntity
							.ok(((GridCellException) ex).getMessage() + ((GridCellException) ex).getCode());
				}
			}

			String fileUrl = null;
			if (GridJsWorkbook.CacheImp != null) {
				fileUrl = GridJsWorkbook.CacheImp.getFileUrl(uid + "/" + filename);
			} else {
				fileUrl = "/GridJs2/GetFile?id=" + uid + "&filename=" + filename;
			}
			return ResponseEntity.ok("\""+fileUrl+"\"");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			return new ResponseEntity<>(gwb.errorJson(e.getMessage()), HttpStatus.OK);
		}
	}

 
}
