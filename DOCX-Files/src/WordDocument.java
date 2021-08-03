import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

public class WordDocument {

	static Connection conn = null;

	static String workingDir = "";
	static String tempFile_filepath = "";

	public static void main(String[] args)
			throws InvalidFormatException, IOException, XmlException, InterruptedException {

		workingDir = getCurrentDir();
		tempFile_filepath = workingDir + "\\src\\Resources\\Temporary.docx";

		ArrayList<sqlRow> data = extractSQL();

		String templatePath = workingDir + "\\src\\Resources\\Template.docx";
		String outputPath = workingDir + "\\src\\OutputDocuments\\Output.docx";
		generateDocument(templatePath, outputPath, data);
	}

	public static ArrayList<sqlRow> extractSQL() {

		ArrayList<sqlRow> tempArr = new ArrayList<sqlRow>();

		conn = sqlManager.openConnection();
		System.out.println("Connected to database");
		try {

			String SQL = "SELECT description, quantity, unit_price, unit_price * quantity as itemCost FROM tblInvoiceDetails WHERE invoice_id = 3";
			Statement stmt = null;
			ResultSet rs = null;

			stmt = conn.createStatement();
			rs = stmt.executeQuery(SQL);

			while (rs.next()) {
				sqlRow tempRow = new sqlRow(rs.getString(1), rs.getString(2), rs.getString(3), rs.getString(4));
				tempArr.add(tempRow);
			}

		} catch (SQLException e) {
			System.out.println("SQLException");
			e.printStackTrace();
		}

//		for (sqlRow row : tempArr) {
//			System.out.println(row.toString());
//
//		}
		return tempArr;
	}

	public static void generateDocument(String source, String destination, ArrayList<sqlRow> data)
			throws InvalidFormatException, IOException, XmlException {

		XWPFDocument doc = new XWPFDocument(OPCPackage.open(source));
		doc = resizeDocumentTable(doc, data.size());
		saveDocument(doc, tempFile_filepath);

		XWPFDocument doc2 = new XWPFDocument(OPCPackage.open(tempFile_filepath));
		doc2 = insertIntoTable(doc2, data);
		String finalDestination = saveDocument(doc2, destination);

		File temp = new File(tempFile_filepath);
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

	public static String saveDocument(XWPFDocument document, String destination) throws IOException {

		String savingDestination = destination;

		if (!savingDestination.equals(tempFile_filepath)) {
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

		FileOutputStream out = null;
		try {
			out = new FileOutputStream(savingDestination);
			document.write(out);
			out.close();

		} catch (IOException e) {
			System.out.println("Error saving file");
			e.printStackTrace();
		} finally {
			out.close();
		}
		return savingDestination;
	}

	public static String getCurrentDir() {
		Path currentRelativePath = Paths.get("");
		return currentRelativePath.toAbsolutePath().toString();
	}

}
