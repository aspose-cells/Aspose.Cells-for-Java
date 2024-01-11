package com.aspose.gridweb.test;

import java.io.IOException;
import java.io.PrintWriter;
import java.lang.reflect.Method;
import java.net.InetAddress;

import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.GridWorksheetCollection;
import com.aspose.gridweb.Unit;

import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
//
//import com.aspose.cells.GridWebBean;
//import com.aspose.cells.GridWorksheet;
//import com.aspose.cells.GridWorksheetCollection;
//import com.aspose.cells.Unit;
//import com.aspose.cells

/**
 * Servlet implementation class TestGridWebServlet
 */
public class TestGridWebServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;
	private static GridWebBean gridweb = null;
	protected PrintWriter out = null;
	private String path =  null;

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doGet(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		this.doPost(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doPost(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {
		try {
			request.setCharacterEncoding("utf-8");

			// 初始化gridWeb
			if (gridweb == null) {
				gridweb = new GridWebBean("hello_bean");
			}
			out = response.getWriter();
			path = request.getSession().getServletContext().getRealPath("/");//.getServletContext().getRealPath("/");
			 
			
			gridweb.setACWLanguageFileUrl("grid/acw_client/lang_en.js");
			out.print(gridweb.getHTMLHead());
			// 做相关操作
			process(request, response);
			 
			gridweb.prepareRender();
			out.print( gridweb.getHTMLBody());
		} catch (Exception ex) {
			throw new ServletException(ex);
		}
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void process(HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		String action = request.getParameter("flag");// 此时servlet的方法执行交给了前端请求控制，根据flag的值，确定需要执行的方法
		try {
			Class clz = this.getClass();// 执行反射即可
			Method method = clz.getDeclaredMethod(action,
					HttpServletRequest.class, HttpServletResponse.class);
			method.invoke(this, request, response);
		} catch (Exception ex) {
			throw ex;
		}
	}

	// 默认的Reload data
	public void reload(HttpServletRequest request, HttpServletResponse response)
			throws Exception {

	 
		gridweb.setWidth(Unit.Pixel(1000));
		gridweb.setHeight(Unit.Pixel(400));

		gridweb.importExcelFile(path + "file\\data.xlsx");
		System.out.println("get user ip is:"+getUserRealIP(request));

	}

	//Add
	public void add(HttpServletRequest request, HttpServletResponse response) {

		GridWorksheetCollection gridWorksheetCollection = gridweb
				.getWorkSheets();
		int i= (gridWorksheetCollection.getCount() );
		GridWorksheet gw=gridWorksheetCollection.add("Sheet"+ i);
		gridweb.setActiveSheetIndex(gw.getIndex());
	}

	//Add Copy
	public void copy(HttpServletRequest request, HttpServletResponse response) throws Exception {
		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		System.err.println(gridweb.getActiveSheetIndex());
		gridWorksheetCollection.addCopy(gridweb.getActiveSheetIndex());
	}

	//Remove Active Sheet
	public void remove(HttpServletRequest request, HttpServletResponse response) throws Exception {
		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		gridWorksheetCollection.removeAt(gridweb.getActiveSheetIndex());
	}
	
	public void changesheet(HttpServletRequest request, HttpServletResponse response) throws Exception {
	    int i=gridweb.getActiveSheetIndex();
		gridweb.setActiveSheetIndex(i+1);
	}
	
	
	public void style1(HttpServletRequest request, HttpServletResponse response) {

		gridweb.setPresetStyle(1);
		System.out.println(gridweb.getFrameTableStyle().getTopBorderStyle().getBorderColor().toString()+",color1 expected:#BB8855");
	}
	public void style2(HttpServletRequest request, HttpServletResponse response) {

		gridweb.setPresetStyle(2);
		System.out.println(gridweb.getFrameTableStyle().getTopBorderStyle().getBorderColor().toString()+",color2 expected:#3366cc");
	}
	public void custstyle2(HttpServletRequest request, HttpServletResponse response) {

		//gridweb.setPresetStyle(7);
		gridweb.setCustomStyleFileName("http://localhost:7080/simple_porting_web/xml/CustomStyle2.xml");
		System.out.println(gridweb.getFrameTableStyle().getTopBorderStyle().getBorderColor().toString()+",cust2 expected:#C0FFC0");
	}
	public void custstyle1(HttpServletRequest request, HttpServletResponse response) {

		//gridweb.setPresetStyle(7);
		gridweb.setCustomStyleFileName("http://localhost:7080/simple_porting_web/xml/CustomStyle1.xml");
		System.out.println(gridweb.getFrameTableStyle().getTopBorderStyle().getBorderColor().toString()+",cust1 expected:#FF6347");
	}
	
public static String getUserRealIP(HttpServletRequest request) throws Exception {
        
        String ip = "";
        
        // 有的user可能使用代理，为处理用户使用代理的情况，使用x-forwarded-for
        if  (request.getHeader("x-forwarded-for") == null)  {
            ip = request.getRemoteAddr();
        }  else  {
            ip = request.getHeader("x-forwarded-for");
        }
        
        if  ("127.0.0.1".equals(ip))  {
            // 获取本机真正的ip地址
            ip = InetAddress.getLocalHost().getHostAddress();
        }
        
        return ip;
    } 

}
