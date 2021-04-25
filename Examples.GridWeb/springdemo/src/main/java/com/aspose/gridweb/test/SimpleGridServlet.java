package com.aspose.gridweb.test;

import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.gridweb.ExtPage;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.Unit;

/**
 * Servlet implementation class SimpleGridServlet
 */
public class SimpleGridServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;
	protected   ExtPage page =ExtPage.getInstance();  
	protected String path = null;
	protected PrintWriter out = null;
	String filename = "D:\\codebase\\local\\grid\\java_version\\workspace\\meetdev\\testweb\\WebContent\\file\\list.xls";
    /**
     * @see HttpServlet#HttpServlet()
     */
    public SimpleGridServlet() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doPost(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		try {
			out = response.getWriter();
			page.setServlet(request,response);
		GridWebBean  gridweb=page.getBean();
		 
		try {
			request.setCharacterEncoding("UTF-8");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		response.setCharacterEncoding("UTF-8");
		path = request.getServletContext().getRealPath("/");
		//webPath = request.getServletContext().getContextPath();

		
			
			
			gridweb.setWidth(Unit.Pixel(800));
			gridweb.setHeight(Unit.Pixel(400));
			//String filename = null;
			path = request.getServletContext().getRealPath("/");
			try {
				gridweb.importExcelFile(  filename);
			} catch (Exception e) {
				e.printStackTrace();
			}

		gridweb.prepareRender();
			String html = gridweb.getHTMLBody();
			out.print(html);
//			FileUtil.putFile(html);

			out.flush();

		} catch (Exception e) {
			e.printStackTrace();
			out.print(e.getMessage());
			out.flush();
		} finally {
			out.close();
		}
	}

}
