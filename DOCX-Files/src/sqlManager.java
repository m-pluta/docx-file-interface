import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class sqlManager {
	// The default connection details for the database
	static String DEFAULT_url = "jdbc:mysql://localhost:3306/dbNEA?serverTimezone=GMT";
	static String DEFAULT_username = "root";
	static String DEFAULT_password = "root";

	// Opens a connection to the database
	public static Connection openConnection(String url, String username, String password) {
		Connection conn = null;
		try {
			Class.forName("com.mysql.cj.jdbc.Driver");
			conn = DriverManager.getConnection(url, username, password);
		} catch (ClassNotFoundException cE) {
			cE.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return conn;
	}

	// Opens connection to db, no specific url given so the default url is used.
	public static Connection openConnection(String username, String password) {
		return openConnection(DEFAULT_url, username, password);
	}

	// Opens connection to db, no specific url, username or password given so a
	// default connection is opened
	public static Connection openConnection() {
		return openConnection(DEFAULT_url, DEFAULT_username, DEFAULT_password);
	}

	// Closes the connection to the database
	public static boolean closeConnection(Connection conn) {
		try {
			if (!conn.isClosed()) { // Checks if the connection is already closed to prevent an exception from
									// happening
				conn.close();
				return true;
			}
		} catch (SQLException e) {
			System.out.println("Could not close!");
			e.printStackTrace();
		}
		return false;
	}

}
