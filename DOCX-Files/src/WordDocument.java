import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

public class WordDocument {

	static String workingDir = "";
	static String tempFile_filepath = "";

	public static void main(String[] args)
			throws InvalidFormatException, IOException, XmlException, InterruptedException {

		workingDir = getCurrentDir();
		tempFile_filepath = "\\src\\Resources\\Temporary.docx";

		ArrayList<sqlRow> data = extractSQL();
		String templatePath = workingDir + "\\src\\Resources\\Template.docx";

		generateDocument(templatePath, workingDir + "\\src\\OutputDocuments\\Output.docx", data);

	}

	public static ArrayList<sqlRow> extractSQL() {

		ArrayList<String> authData = getAuthData(workingDir + "\\src\\authDB.json");
		ArrayList<sqlRow> tempArr = new ArrayList<sqlRow>();

		try {
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection conn = DriverManager.getConnection(authData.get(0), authData.get(1), authData.get(2));
			System.out.println("Connected to database");

			String SQL = "SELECT * FROM " + authData.get(3);
			Statement statement;
			ResultSet result;

			statement = conn.createStatement();
			result = statement.executeQuery(SQL);

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

	static ArrayList<String> getAuthData(String source) {

		ArrayList<String> tempArray = new ArrayList<String>();
		JSONParser parser = new JSONParser();

		try {

			JSONObject obj = (JSONObject) parser.parse(new FileReader(source));
			tempArray.add((String) obj.get("url"));
			tempArray.add((String) obj.get("username"));
			tempArray.add((String) obj.get("password"));
			tempArray.add((String) obj.get("tableName"));

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (ParseException e) {
			e.printStackTrace();
		}

		return tempArray;

	}

	public static void generateDocument(String source, String destination, ArrayList<sqlRow> data)
			throws InvalidFormatException, IOException, XmlException, InterruptedException {

		XWPFDocument doc = new XWPFDocument(OPCPackage.open(source));
		doc = resizeDocumentTable(doc, data.size());
		saveDocument(doc, workingDir + tempFile_filepath);

		XWPFDocument doc2 = new XWPFDocument(OPCPackage.open(workingDir + tempFile_filepath));

		doc2 = insertIntoTable(doc2, data);
		String finalDestination = saveDocument(doc2, destination);

		File temp = new File(workingDir + tempFile_filepath);
//		setHiddenAttribute(temp, false);
		if (temp.delete()) {
			System.out.println("Temporary file deleted successfully");
		} else {
			System.out.println("Failed to delete file");
		}

		System.out.println("Saved document under: " + finalDestination);

	}

	public static XWPFDocument insertIntoTable(XWPFDocument document, ArrayList<sqlRow> data) {

		List<XWPFTable> tables = document.getTables();

		XWPFTable table = tables.get(0);

		for (int i = 1; i < data.size() + 1; i++) {
			int j = 0;
			for (XWPFTableCell cell : table.getRows().get(i).getTableCells()) {
				for (XWPFParagraph p : cell.getParagraphs()) {
					for (XWPFRun r : p.getRuns()) {
						r.setText(data.get(i - 1).data[j++]);
					}
				}
			}
			j = 0;
		}

		return document;
	}

	public static XWPFDocument resizeDocumentTable(XWPFDocument document, int amtRows)
			throws XmlException, IOException {

		List<XWPFTable> tables = document.getTables();
		XWPFTable table = tables.get(0);

		XWPFTableRow blankRow = table.getRows().get(2);
		if (amtRows == 0) {
			table.removeRow(2);
			table.removeRow(1);

		} else if (amtRows == 1) {
			table.removeRow(2);

		} else if (amtRows > 2) {
			int newRowsNeeded = amtRows - 2;
			int startRowInsert = 2;
			for (int i = 2; i < newRowsNeeded + 2; i++) {
				CTRow ctrow = CTRow.Factory.parse(blankRow.getCtRow().newInputStream());
				XWPFTableRow newRow = new XWPFTableRow(ctrow, table);

				for (XWPFTableCell cell : newRow.getTableCells()) {
					for (XWPFParagraph paragraph : cell.getParagraphs()) {
						for (XWPFRun run : paragraph.getRuns()) {
							run.setText("", 0);
						}
					}
				}
				table.addRow(newRow, startRowInsert++);
			}
		} else {
			System.out.println("Not a valid amount of rows specified");
		}

		return document;
	}

	public static String saveDocument(XWPFDocument document, String destination) throws InterruptedException {

		String savingDestination = destination;

		if (!savingDestination.equals(workingDir + tempFile_filepath)) {
			File f = new File(savingDestination);
			if (f.exists() && !f.isDirectory()) {
				System.out.println("File already exists under " + f.toPath().toString());

				boolean found = false;
				int counter = 1;
				while (!found) {
					File t = new File(workingDir + "\\src\\OutputDocuments\\Output" + counter + ".docx");

					if (t.exists() && !t.isDirectory()) {
						counter++;
					} else {
						savingDestination = workingDir + "\\src\\OutputDocuments\\Output" + counter + ".docx";
						found = true;
					}
				}
			}

		}

		try {
			document.write(new FileOutputStream(savingDestination));

		} catch (IOException e) {
			System.out.println("Error saving file");
			e.printStackTrace();
		}
		return savingDestination;
	}

	public static String getCurrentDir() {
		Path currentRelativePath = Paths.get("");
		return currentRelativePath.toAbsolutePath().toString();
	}

}
