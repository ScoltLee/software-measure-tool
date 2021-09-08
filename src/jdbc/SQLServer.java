package jdbc;


import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.sql.ResultSet;
import javax.naming.spi.DirStateFactory.Result;

public class SQLServer {

	static String Url= "jdbc:mysql://localhost:3306/QLShowroom"; 
	static String usr=	"SCOLTLEE";
	static String pass=	"";
	
	
	
 
    
    public static void main(String args[]) {
        try {
            // connnect to database 'testdb'
            Connection conn = getConnection(Url, usr, pass);
            // crate statement
            Statement stmt = conn.createStatement();
            // get data from table 'student'
            ResultSet rs = stmt.executeQuery("select * from Xe");
            // show data
            
            // close connection
            conn.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
 
 
   
    public static Connection getConnection(String dbURL, String userName, 
            String password) {
        Connection conn = null;
        try {
            Class.forName("com.mysql.jdbc.Driver");
            conn = DriverManager.getConnection(dbURL, userName, password);
            System.out.println("connect successfully!");
        } catch (Exception ex) {
            System.out.println("connect failure!");
           ex.printStackTrace(); 
      
        }
        return conn;
    }
}