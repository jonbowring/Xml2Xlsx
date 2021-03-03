package xml2xlsx.jar;

import javax.xml.parsers.DocumentBuilder; 
import javax.xml.parsers.DocumentBuilderFactory;
import org.xml.sax.ErrorHandler;
import org.xml.sax.SAXException; 
import org.xml.sax.SAXParseException;
import org.xml.sax.helpers.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.DocumentType;
import org.w3c.dom.Element;
import org.w3c.dom.Entity;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class AppXml2Xlsx {

	static final String outputEncoding = "UTF-8";
	
	
	public static void main(String[] args) throws Exception {

		// Declare and initialise the variables
		String src = "", 
				tgt = "";
		
		// TODO get the command line params
		// src
		src = "C:\\Users\\jbowring\\OneDrive - Informatica\\_My Documents\\Coding\\Windows\\Excel\\Xml2Xlsx\\books.xml";
		// tgt
		tgt = "C:\\Users\\jbowring\\OneDrive - Informatica\\_My Documents\\Coding\\Windows\\Excel\\Xml2Xlsx\\out.xlsx";
		
		// Parse the source XML file
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder(); 
		Document doc = db.parse(new File(src));
		dbf.setNamespaceAware(true);
		
		// Initialise the target Excel workbook
		Workbook xlWorkbook = new XSSFWorkbook();
		CreationHelper createHelper = xlWorkbook.getCreationHelper();
		
		// Get the workbook node
		Element workbook = (Element) doc.getElementsByTagName("workbook").item(0);
		System.out.println(workbook.getNodeName());
		
		// Get all worksheets in the workbook and loop trhough them
		NodeList worksheets = workbook.getElementsByTagName("worksheet");
		for(int s = 0; s < worksheets.getLength(); s++) {
			
			// Get the current worksheet
			Element worksheet = (Element) worksheets.item(s);
			String sheetName = worksheet.getAttribute("name");
			System.out.println("\t" + worksheet.getNodeName());
			
			// Initialise the target Excel worksheet
			Sheet xlSheet = xlWorkbook.createSheet(sheetName);
			
			// Get all rows in the current worksheet and loop through them
			NodeList rows = worksheet.getElementsByTagName("row");
			for(int r = 0; r < rows.getLength(); r++) {
				
				// Get the current row
				Element row = (Element) rows.item(r);
				System.out.println("\t\t" + row.getNodeName());
				
				// Initialise the target row
				Row xlRow = xlSheet.createRow(r);
				
				// Get all cells in the current row and loop through them
				NodeList cells = row.getElementsByTagName("cell");
				for(int c = 0; c < cells.getLength(); c++) {
					
					// Get the current cell
					Element cell = (Element) cells.item(c);
					String cellValue = cell.getTextContent();
					System.out.println("\t\t\t" + cell.getTextContent());
					
					// Initialise the target Excel cell and add the value
					Cell xlCell = xlRow.createCell(c);
					xlCell.setCellValue(cellValue);
					
				}
				
			}
			
		}
			
		// Save and close the target Excel file
		try (OutputStream fileOut = new FileOutputStream(tgt)) {
		    xlWorkbook.write(fileOut);
		}

	}

}
