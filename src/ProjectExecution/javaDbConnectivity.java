package ProjectExecution;
import java.sql.*;
public class javaDbConnectivity {

	public static void main(String[] args) {
		try{
		Class.forName("oracle.jdbc.driver.OracleDriver");
		Connection con=DriverManager.getConnection(
				"jdbc:oracle:thin:@10.3.64.78:1521:xe","iAFDB","iAFDB");
		Statement stmt=con.createStatement();
		ResultSet rs=stmt.executeQuery("select * from ASSIGNED_MAC where RUN_ID='41327'");
		while(rs.next()){
			System.out.println(rs.getInt(4)+" "+rs.getString(3)+" "+rs.getString(4)+" "+rs.getString(5));
			}
		rs.close();
		stmt.close();
		con.close();

		}catch(Exception e){
		 System.out.println(e);
		}
	}

}
