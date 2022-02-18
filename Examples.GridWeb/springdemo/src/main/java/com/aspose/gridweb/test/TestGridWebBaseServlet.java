package com.aspose.gridweb.test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.core.io.ClassPathResource;
//
//import com.aspose.cells.ExtPage;
//import com.aspose.cells.GridWebBean;
//import com.aspose.cells.ManualLog;
//import com.aspose.cells.Unit;
import com.aspose.gridweb.ExtPage;
import com.aspose.gridweb.GridWebBean;
//import com.aspose.gridweb.ManualLog;
import com.aspose.gridweb.Unit;

public abstract class TestGridWebBaseServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;
	protected   ExtPage page =ExtPage.getInstance();
	protected PrintWriter out = null;
	protected String path = null;
	protected String webPath = null;
	
	public TestGridWebBaseServlet() {
		super();
 
	}
    
	protected void doGet(HttpServletRequest request, HttpServletResponse response) {

		doPost(request, response);
	}

	protected void doPost(HttpServletRequest request, HttpServletResponse response) {
		try {
  
			 request.setCharacterEncoding("utf-8"); 
			  response.setContentType("text/html;charset=utf-8"); 
			  response.setCharacterEncoding("utf-8"); 
			    out = response.getWriter(); //在设置完编码以后在获取输出流就好了。 
		page.setServlet(request,response);
		GridWebBean  gridweb=page.getBean();
		gridweb.setWidth(Unit.Pixel(800));
		gridweb.setHeight(Unit.Pixel(400));
 
		path = request.getServletContext().getRealPath("/");
		webPath = request.getServletContext().getContextPath();

		
			
			
			
		// do the reflect method
			this.process(gridweb,request, response);

		gridweb.prepareRender();
			String html = gridweb.getHTMLBody();
			out.print(html);
//			FileUtil.putFile(html);

//			out.flush();

		} catch (IOException e) {
			e.printStackTrace();

		} catch (Exception e) {
			e.printStackTrace();
			out.print(e.getMessage());
//			out.flush();

		} finally {
//			out.close();
		}
		
	}

	@SuppressWarnings("unchecked")
	public void process(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		String action = request.getParameter("flag");
		if (action == null) {
			return;
		}

		@SuppressWarnings("rawtypes")
		Class clz = this.getClass();
		Method method = null;
		try {
			method = clz.getDeclaredMethod(action,GridWebBean.class, HttpServletRequest.class, HttpServletResponse.class);
			method.invoke(this,gridweb, request, response);
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		}  
	}

	// the default Reload data
	protected void reloadfile(GridWebBean gridweb,HttpServletRequest request, String file) {

		 

		gridweb.setWidth(Unit.Pixel(800));
		gridweb.setHeight(Unit.Pixel(400));
		String filename = null;
		path = request.getServletContext().getRealPath("/");
		
		//spring way
		ClassPathResource cpr = new ClassPathResource("file"+File.separator+file);
		
		try {
//			gridweb.importExcelFile(path + "file\\" + file);
			InputStream in = cpr.getInputStream();
			gridweb.importExcelFile(in);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public abstract void reload(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response);

}
