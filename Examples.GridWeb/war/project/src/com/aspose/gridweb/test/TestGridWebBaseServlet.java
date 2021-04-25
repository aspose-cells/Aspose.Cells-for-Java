package com.aspose.gridweb.test;

import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.Serializable;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.gridweb.ExtPage;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.PresetStyle;
import com.aspose.gridweb.Unit;

public abstract class TestGridWebBaseServlet extends HttpServlet{
	private static final long serialVersionUID = 1L;
	protected   ExtPage page =ExtPage.getInstance();
	protected PrintWriter out = null;
	protected String path = null;
	protected String webPath = null;
	public TestGridWebBaseServlet() {
		super();
		ExtPage.setMaxholders(1000);
		ExtPage.setMemoryinstanceexpires(60);
		ExtPage.setMemoryCleanRateTime(120);
		ExtPage.setTempfilepath("c:/temp/");
	}
	 
	protected void doGet(HttpServletRequest request, HttpServletResponse response) {

		doPost(request, response);
		}

	protected void doPost(HttpServletRequest request, HttpServletResponse response) {
		try {
			response.setCharacterEncoding("UTF-8");
//			response.setHeader("content-type","text/html;charset=UTF-8");
			out = response.getWriter();
		//may throw exception 
		page.setServlet(request, response);
		GridWebBean  gridweb=page.getBean();
		//we shall call it to update request and response in gridweb before render
		 
		try {
			request.setCharacterEncoding("UTF-8");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		
		path = request.getServletContext().getRealPath("/");
		webPath = request.getServletContext().getContextPath();

		 
			
			
			// do the reflect method
			this.process(gridweb,request, response);

			gridweb.prepareRender();
			String html = gridweb.getHTMLBody();
			out.print(html);
//			FileUtil.putFile(html);

			out.flush();
			
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e2) {
			//e.printStackTrace();
			out.print(e2.getMessage());
			out.flush();

		}finally{
			out.close();
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
		try {
			gridweb.importExcelFile(path + "file\\" + file);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public abstract void reload(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response);

}
