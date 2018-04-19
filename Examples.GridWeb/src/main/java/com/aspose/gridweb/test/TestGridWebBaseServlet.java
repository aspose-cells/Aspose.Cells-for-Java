package com.aspose.gridweb.test;

import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.gridweb.ExtPage;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.Unit;
import com.aspose.gridweb.test.util.FileUtil;

public abstract class TestGridWebBaseServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;
	protected   ExtPage page =ExtPage.getInstance();
	protected PrintWriter out = null;
	protected String path = null;
	protected String webPath = null;

	 
	protected void doGet(HttpServletRequest request, HttpServletResponse response) {

		doPost(request, response);
		}

	protected void doPost(HttpServletRequest request, HttpServletResponse response) {

		GridWebBean  gridweb=page.getBean(request);
		//we shall call it to update request and response in gridweb before render
		gridweb.setReqRes(request, response);
		try {
			request.setCharacterEncoding("UTF-8");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		response.setCharacterEncoding("UTF-8");
		path = request.getServletContext().getRealPath("/");
		webPath = request.getServletContext().getContextPath();

		try {
			out = response.getWriter();
			
			
			// do the reflect method
			this.process(gridweb,request, response);

			gridweb.prepareRender();
			String html = gridweb.getHTMLBody();
			out.print(html);
//			FileUtil.putFile(html);

			out.flush();
			
		} catch (IOException e) {
			e.printStackTrace();
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
