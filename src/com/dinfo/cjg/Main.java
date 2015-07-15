package com.dinfo.cjg;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.dbutils.BeanProcessor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Main {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
//		String path=System.getProperty("user.dir");
//		ArrayList<List<String>> list =new ArrayList<>();
//		ExcelTools.readxml(list,path+"/model.xml");
//		
//		HSSFWorkbook book=ExcelTools.downWorkbook(list);
//		
		//System.err.println(path);
	    int  key=1; // 1.证券-基金-期货   2.虚假陈述-内幕交易-关联交易   3.证券公司-基金公司-期货公司
		String sql="select *  from result3";
		Connection conn=getConnection.getCon();
		Statement st=null;
		ResultSet rs=null;
		ArrayList<area_moneyDto> list=null;
	  try {
		  st=conn.createStatement();
		  rs=st.executeQuery(sql);
		     area_moneyDto dto=new area_moneyDto();  
			  BeanProcessor pro = new BeanProcessor();
		       list =(ArrayList<area_moneyDto>)pro.toBeanList(rs,dto.getClass());
	     } catch (Exception e) {
		// TODO: handle exception
	   }
	  finally{
		  try {
			rs.close();
			 st.close();
			  conn.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 
	  }
	  HSSFWorkbook book=ExcelTools.writeWorkbook(list,key);
	  OutputStream out;
		try {
			out = new FileOutputStream("d://4-24.xls");
			book.write(out);
		}
		catch (Exception e) {
			// TODO: handle exception
		}
	}
	
}
