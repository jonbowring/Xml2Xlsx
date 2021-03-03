package xml2xlsx.jar;

import javax.xml.parsers.DocumentBuilder; 
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class AppXml2Xlsx {

	static final String outputEncoding = "UTF-8";
	
	
	public static void main(String[] args) throws Exception {

		// Declare and initialise the variables
		String src = "", 
				tgt = "";
		
		// Parse the command line arguments
		for(int p = 0; p < args.length; p++) {
			switch(args[p]) {
				case "--src":
					src = args[p + 1];
					p++;
					break;
				case "--tgt":
					tgt = args[p + 1];
					p++;
					break;
				default:
					break;
			}
		}
		
		// Parse the source XML file
		System.out.println("Reading XML source file '" + src + "'...");
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder(); 
		Document doc = db.parse(new File(src));
		dbf.setNamespaceAware(true);
		
		// Initialise the target Excel workbook
		Workbook xlWorkbook = new XSSFWorkbook();
		CreationHelper xlHelper = xlWorkbook.getCreationHelper();
		
		// Get the workbook node
		Element workbook = (Element) doc.getElementsByTagName("workbook").item(0);
		System.out.println("Initialising workbook...");
		
		// Get all worksheets in the workbook and loop trhough them
		NodeList worksheets = workbook.getElementsByTagName("worksheet");
		for(int s = 0; s < worksheets.getLength(); s++) {
			
			// Get the current worksheet
			Element worksheet = (Element) worksheets.item(s);
			String sheetName = worksheet.getAttribute("name");
			System.out.println("Adding worksheet '" + sheetName + "'...");
			
			// Initialise the target Excel worksheet
			Sheet xlSheet = xlWorkbook.createSheet(sheetName);
			
			// Get all rows in the current worksheet and loop through them
			NodeList rows = worksheet.getElementsByTagName("row");
			for(int r = 0; r < rows.getLength(); r++) {
				
				// Get the current row
				Element row = (Element) rows.item(r);
				
				// Initialise the target row
				Row xlRow = xlSheet.createRow(r);
				
				// Get all cells in the current row and loop through them
				NodeList cells = row.getElementsByTagName("cell");
				for(int c = 0; c < cells.getLength(); c++) {
					
					// Get the current cell
					Element cell = (Element) cells.item(c);
					String cellValue = cell.getTextContent();
					String cellType = cell.getAttribute("type");
					
					
					// Initialise the target Excel cell and add the value
					Cell xlCell = xlRow.createCell(c);
					switch(cellType) {
						case "formula":
							xlCell.setCellFormula(cellValue);
							break;
						case "string":
							xlCell.setCellValue(cellValue);
							break;
						case "int":
							int cellInt = Integer.parseInt(cellValue);
							xlCell.setCellValue(cellInt);
							break;
						case "float":
							Double cellDouble = Double.parseDouble(cellValue);
							xlCell.setCellValue(cellDouble);
							break;
						case "date":
							SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
							Date cellDate = fmt.parse(cellValue);
							xlCell.setCellValue(cellDate);
							CellStyle style = xlWorkbook.createCellStyle();
							style.setDataFormat(xlHelper.createDataFormat().getFormat("dd/mm/yyyy hh:mm:ss"));
							xlCell.setCellStyle(style);
							break;
					}
					
				}
				
			}
			
		}
			
		// Save and close the target Excel file
		try (OutputStream fileOut = new FileOutputStream(tgt)) {
		    System.out.println("Saving Excel file '" + tgt + "'...");
			xlWorkbook.write(fileOut);
		    xlWorkbook.close();
		    System.out.println("File saved!");
		}

	}

}
