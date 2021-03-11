import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class WordDocument {

	static String workingDir = "";

	public static void main(String[] args) {

		ArrayList<sqlRow> data = extractSQL("jdbc:mysql://freedb.tech:3306/freedbtech_dbMikey",
				"freedbtech_ThunderCandy", "password", "tblPizza");

//		for (sqlRow node : data) { // Debug
//			System.out.format("%s, %s, %s, %s \n", node.get(0), node.get(1), node.get(2), node.get(3));
//		}

		workingDir = getCurrentDir();

		String templatePath = workingDir + "\\src\\Resources\\Template.docx";

	}

	public static ArrayList<sqlRow> extractSQL(String url, String username, String password, String tableName) {

		ArrayList<sqlRow> tempArr = new ArrayList<sqlRow>();

		try {
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection conn = DriverManager.getConnection(url, username, password);
			System.out.println("Connected to database");

			String SQL = "SELECT * FROM " + tableName;
			Statement statement;
			ResultSet result;

			statement = conn.createStatement();
			result = statement.executeQuery(SQL);

			boolean first = true;
			while (result.next()) {
				sqlRow tempRow = new sqlRow(result.getString("PizzaType"), result.getString("Toppings"),
						result.getString("Cost"), result.getString("SpecialInfo"));
				tempArr.add(tempRow);
			}

		} catch (SQLException e) {
			System.out.println("SQLException: " + e.toString());
		} catch (ClassNotFoundException e) {
			System.out.println("Class not found exception: " + e.toString());
		}

		return tempArr;
	}

	public static String getCurrentDir() {
		Path currentRelativePath = Paths.get("");
		return currentRelativePath.toAbsolutePath().toString();
	}

}
