package com.informatica.xml2xlsx;

import javax.xml.parsers.DocumentBuilder; 
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import java.util.HashMap;

public class AppXml2Xlsx {

	static final String outputEncoding = "UTF-8";
	
	
	public static void main(String[] args) throws Exception {

		// Declare and initialise the variables
		String src = "", 
				tgt = "";
		HashMap<String, Style> styleMap = new HashMap<String, Style>();
		
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
		XPathFactory xpathFactory = XPathFactory.newInstance();
		XPath xpath = xpathFactory.newXPath();
		Document doc = db.parse(new File(src));
		dbf.setNamespaceAware(true);
		
		// Initialise the target Excel workbook
		Workbook xlWorkbook = new XSSFWorkbook();
		CreationHelper xlHelper = xlWorkbook.getCreationHelper();
		
		// Get the workbook node
		Element workbook = (Element) xpath.evaluate("/workbook", doc, XPathConstants.NODE);
		System.out.println("Initialising workbook...");
		
		// Parse the styles
		NodeList styles = (NodeList) xpath.evaluate("/workbook/styles/style", doc, XPathConstants.NODESET);
		
		// If a styles array has been included
		if(styles.getLength() > 0) {
			
			// Loop through the styles and create the style objects
			for(int y = 0; y < styles.getLength(); y++) {
				
				// Initialise the style object
				Element styleEl = (Element) styles.item(y);
				Style style = new Style(styleEl.getAttribute("name"));
				
				// If it has the valign attribute then save it
				if(styleEl.hasAttribute("valign")) {
					style.setVAlign(styleEl.getAttribute("valign"));
				}
				
				// If it has the halign attribute then save it
				if(styleEl.hasAttribute("halign")) {
					style.setHAlign(styleEl.getAttribute("halign"));
				}
				
				// Add the current style to the hash map
				styleMap.put(styleEl.getAttribute("name"), style);
			}
		}
		
		
		// Get all worksheets in the workbook and loop trhough them
		NodeList worksheets = (NodeList) xpath.evaluate("/workbook/worksheet", doc, XPathConstants.NODESET);
		//NodeList worksheets = workbook.getElementsByTagName("worksheet");
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
					
					// Get the cell type if it is available
					String cellType = "";
					if(cell.hasAttribute("type")) {
						cellType = cell.getAttribute("type");
					}
					
					
					
					// Initialise the target Excel cell and add the value
					Cell xlCell = xlRow.createCell(c);
					CellStyle cellStyle = xlWorkbook.createCellStyle();
					
					// Apply the cell data types
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
							cellStyle.setDataFormat(xlHelper.createDataFormat().getFormat("dd/mm/yyyy hh:mm:ss"));
							break;
						default:
							xlCell.setCellValue(cellValue);
							break;
					}
					
					// If a cell style has been applied then add it to the cell
					if(cell.hasAttribute("style")) {
						
						// Get the style settings
						Style style = styleMap.get(cell.getAttribute("style"));
						
						// Apply the vertical alignment if set
						String valign = style.getVAlign();
						if(valign.length() > 0) {
							
							switch(valign) {
								case "top":
									cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
									break;
								case "center":
									cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
									break;
								case "bottom":
									cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
									break;
								default:
									break;
							}
							
						}
						
						// Apply the horizontal alignment if set
						String halign = style.getHAlign();
						if(halign.length() > 0) {
							
							switch(halign) {
								case "left":
									cellStyle.setAlignment(HorizontalAlignment.LEFT);
									break;
								case "center":
									cellStyle.setAlignment(HorizontalAlignment.CENTER);
									break;
								case "right":
									cellStyle.setAlignment(HorizontalAlignment.RIGHT);
									break;
								default:
									break;
							}
							
						}
						
					}
					
					// Save the style to the cell
					xlCell.setCellStyle(cellStyle);
					
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
