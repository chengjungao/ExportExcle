package com.dinfo.cjg;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;




public class getConnection {
	private static final String MYSQL = "jdbc:mysql://localhost:3306/zqmanage";
	public static Connection getCon()
	{
		Connection conn = null;
		try {
			Class.forName("com.mysql.jdbc.Driver");// 加载驱动
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		}
		try {
			conn = DriverManager.getConnection(MYSQL, "root", "admin");
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return conn;
	}
}
