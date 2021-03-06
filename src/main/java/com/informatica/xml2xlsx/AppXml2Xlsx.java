package com.informatica.xml2xlsx;

import javax.xml.parsers.DocumentBuilder; 
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import java.util.HashMap;
import java.util.List;

public class AppXml2Xlsx {

	static final String outputEncoding = "UTF-8";
	
	
	public static void main(String[] args) throws Exception {

		// Declare and initialise the variables
		String src = "", 
				tgt = "";
		HashMap<String, Style> styleMap = new HashMap<String, Style>();
		HashMap<String, Validation> validationMap = new HashMap<String, Validation>();
		StyleHelper styleHelper = new StyleHelper();
		
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
		XSSFWorkbook xlWorkbook = new XSSFWorkbook();
		CreationHelper xlHelper = xlWorkbook.getCreationHelper();
		
		// Get the workbook node
		System.out.println("Initialising workbook...");
		Element workbook = (Element) xpath.evaluate("/workbook", doc, XPathConstants.NODE);
		
		// Parse the styles
		NodeList styles = (NodeList) xpath.evaluate("/workbook/styles/style", doc, XPathConstants.NODESET);
		
		// If a styles array has been included
		if(styles.getLength() > 0) {
			
			// Loop through the styles and create the style objects
			for(int y = 0; y < styles.getLength(); y++) {
				
				// Initialise the style object
				Element styleEl = (Element) styles.item(y);
				Style style = new Style(styleEl.getAttribute("name"), xlWorkbook);
				
				// If the style has an attribute element
				if(styleEl.getElementsByTagName("align").getLength() > 0 ) {
					
					Element align = (Element) styleEl.getElementsByTagName("align").item(0);
					
					// If it has the valign attribute then save it
					if(align.hasAttribute("vertical")) {
						style.setVAlign(align.getAttribute("vertical"));
					}
					
					// If it has the halign attribute then save it
					if(align.hasAttribute("horizontal")) {
						style.setHAlign(align.getAttribute("horizontal"));
					}
					
				}
				
				// If the style has a format element
				if(styleEl.getElementsByTagName("format").getLength() > 0 ) {
					
					Element format = (Element) styleEl.getElementsByTagName("format").item(0);
					
					// If it has the type element then save it 
					if(format.hasAttribute("type")) {
						style.setFormatType(format.getAttribute("type"));
					}
					
					// If it has the pattern element then save it 
					if(format.hasAttribute("pattern")) {
						style.setFormatPattern(format.getAttribute("pattern"));
					}
					
				}
				
				// If the style has a fill element
				if(styleEl.getElementsByTagName("fill").getLength() > 0 ) {
					
					Element fill = (Element) styleEl.getElementsByTagName("fill").item(0);
					
					// If it has the colour element then save it 
					if(fill.hasAttribute("colour")) {
						style.setFillColour(fill.getAttribute("colour"));
					}
					
					// If it has the colour and pattern defined then save it 
					if(fill.hasAttribute("colour") && fill.hasAttribute("pattern")) {
						style.setFillPattern(fill.getAttribute("pattern"));
					}
					// Else if it has the the colour but no pattern then use a default pattern
					else if(fill.hasAttribute("colour") && !fill.hasAttribute("pattern")) {
						style.setFillPattern("solid-foreground");
					}
					
				}
				
				// If the style has a wrap element then set wrap to true
				if(styleEl.getElementsByTagName("wrap").getLength() > 0 ) {
					
					style.setWrap(true);
					
				}
				
				
				// If the style has border elements
				NodeList borders = styleEl.getElementsByTagName("border");
				if(borders.getLength() > 0) {
					
					// Loop through all of the borders
					for(int b = 0; b < borders.getLength(); b++) {
						
						// Initialise the border object
						Element borderEl = (Element) borders.item(b);
						Border border;
						if(borderEl.hasAttribute("type") && borderEl.hasAttribute("colour")) {
							border = new Border(borderEl.getAttribute("pos"), borderEl.getAttribute("type"), borderEl.getAttribute("colour"));
						}
						else if(borderEl.hasAttribute("type")) {
							border = new Border(borderEl.getAttribute("pos"), borderEl.getAttribute("type"));
						}
						else {
							border = new Border(borderEl.getAttribute("pos"));
						}
						
						// Add the border to the style
						style.addBorder(border);
					}
					
				}
				
				// Add the font styles if set
				Element fontEl = (Element) styleEl.getElementsByTagName("font").item(0);
				if(fontEl != null) {
					
					// Initialise the font
					Font font = xlWorkbook.createFont();
					
					// Get the font settings
					String fontName = fontEl.getAttribute("name");
					Element fontSize = (Element) fontEl.getElementsByTagName("size").item(0);
					Element fontItalic = (Element) fontEl.getElementsByTagName("italic").item(0);
					Element fontStrike = (Element) fontEl.getElementsByTagName("strikeout").item(0);
					
					// Set the font name if set
					if(fontName.length() > 0) {
						font.setFontName(fontName);
					}
					
					// Set the font size if set
					if(fontSize != null) {
						font.setFontHeightInPoints(Short.parseShort(fontSize.getTextContent()));
					}
					
					// Set the font italic if set
					if(fontItalic != null) {
						font.setItalic(true);
					}
					
					// Set the font strikeout if set
					if(fontStrike != null) {
						font.setStrikeout(true);
					}
					
					// Save the font
					style.setFont(font);
					
				}
				
				// Add the current style to the hash map
				styleMap.put(styleEl.getAttribute("name"), style);
				
			}
		} // End of if styles.getLength() > 0
		
		// Parse the styles
		NodeList validations = (NodeList) xpath.evaluate("/workbook/validations/validation", doc, XPathConstants.NODESET);
		
		// If a validations array has been included
		if(validations.getLength() > 0) {
			
			// Loop through the validations and create the validation objects
			for(int v = 0; v < validations.getLength(); v++) {
				
				// Initialise the validation object
				Element validationEl = (Element) validations.item(v);
				String validationType = validationEl.getElementsByTagName("type").item(0).getTextContent();
				Validation validation = new Validation(validationEl.getAttribute("name"), validationType);
				
				// If set get and set the formula
				Element formulaEl = (Element) validationEl.getElementsByTagName("formula").item(0);
				if(formulaEl != null) {
					validation.setFormula(formulaEl.getTextContent());
				}
				
				// If set get the list of validation values
				Element valuesArrEl = (Element) validationEl.getElementsByTagName("values").item(0);
				if(valuesArrEl != null) {
					
					NodeList valuesEl = valuesArrEl.getElementsByTagName("value");
					
					if(valuesEl.getLength() > 0) {
						List<String> valuesList = new ArrayList<String>();
						for(int a = 0; a < valuesEl.getLength(); a++) {
							
							Element valueEl = (Element) valuesEl.item(a);
							valuesList.add(valueEl.getTextContent());
							
						}
						
						// Save the list of values to the validation
						validation.setValues(valuesList.toArray(new String[valuesEl.getLength()]));
					}
					
				}
				
				
				
				
				// Save the validation to the map
				validationMap.put(validation.getName(), validation);
			}
			
		}
		
		
		// Get all worksheets in the workbook and loop trhough them
		NodeList worksheets = (NodeList) xpath.evaluate("/workbook/worksheet", doc, XPathConstants.NODESET);
		for(int s = 0; s < worksheets.getLength(); s++) {
			
			// Get the current worksheet
			Element worksheet = (Element) worksheets.item(s);
			String sheetName = worksheet.getAttribute("name");
			System.out.println("Adding worksheet '" + sheetName + "'...");
			
			// Initialise the target Excel worksheet and data validation helper
			XSSFSheet xlSheet = xlWorkbook.createSheet(sheetName);
			XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(xlSheet);
			
			// Get all rows in the current worksheet and loop through them
			NodeList rows = worksheet.getElementsByTagName("row");
			int maxR = 0, maxC = 0;
			for(int r = 0; r < rows.getLength(); r++) {
				
				// Update the max row count
				if(r > maxR) {
					maxR = r;
				}
				
				// Get the current row
				Element row = (Element) rows.item(r);
				
				// Initialise the target row
				XSSFRow xlRow = xlSheet.createRow(r);
				
				// Get all cells in the current row and loop through them
				NodeList cells = row.getElementsByTagName("cell");
				for(int c = 0; c < cells.getLength(); c++) {
					
					// Update the max col count
					if(c > maxC) {
						maxC = c;
					}
					
					// Get the current cell
					Element cell = (Element) cells.item(c);
					String cellValue = cell.getTextContent();
					
					// Initialise the target Excel cell and add the value
					XSSFCell xlCell = xlRow.createCell(c);
					XSSFCellStyle cellStyle = xlWorkbook.createCellStyle();
					
					// If a cell style has been applied then add it to the cell
					if(cell.hasAttribute("style")) {
						
						// Get the style settings
						Style style = styleMap.get(cell.getAttribute("style"));
						
						// Apply the cell format if set
						String formatType = style.getFormatType();
						if(formatType.length() > 0) {
							
							// Apply the cell data types
							switch(formatType) {
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
									SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
									Date cellDate = fmt.parse(cellValue);
									xlCell.setCellValue(cellDate);
									cellStyle.setDataFormat(xlHelper.createDataFormat().getFormat(style.getFormatPattern()));
									break;
								default:
									xlCell.setCellValue(cellValue);
									break;
							}
							
						}
						else {
							xlCell.setCellValue(cellValue);
						}
						
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
						
						// Apply the top border if set
						Border borderTop = style.getBorder("top");
						if(borderTop != null) {
							
							// Apply the style if set
							BorderStyle borderStyle = styleHelper.getBorderStyles().get(borderTop.getType());
							if(borderStyle != null) {
								cellStyle.setBorderTop(borderStyle);
							}
							
							// Apply the colour if set
							IndexedColors borderColour = styleHelper.getColours().get(borderTop.getColour());
							
							// If an rgb colour has been specified then use that
							if(borderTop.getColour().matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {
								
								String[] rgb = borderTop.getColour().substring(4, borderTop.getColour().length() - 1).split("\\s*,\\s*");
								cellStyle.setTopBorderColor(new XSSFColor(new java.awt.Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])), new DefaultIndexedColorMap()));
							}
							else if(borderColour != null) {
								cellStyle.setTopBorderColor(borderColour.getIndex());
							}
							
						}
						
						// Apply the right border if set
						Border borderRight = style.getBorder("right");
						if(borderRight != null) {
							
							// Apply the style if set
							BorderStyle borderStyle = styleHelper.getBorderStyles().get(borderRight.getType());
							if(borderStyle != null) {
								cellStyle.setBorderRight(borderStyle);
							}
							
							// Apply the colour if set
							IndexedColors borderColour = styleHelper.getColours().get(borderRight.getColour());
							
							// If an rgb colour has been specified then use that
							if(borderRight.getColour().matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {
								
								String[] rgb = borderRight.getColour().substring(4, borderRight.getColour().length() - 1).split("\\s*,\\s*");
								cellStyle.setRightBorderColor(new XSSFColor(new java.awt.Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])), new DefaultIndexedColorMap()));
							}
							else if(borderColour != null) {
								cellStyle.setRightBorderColor(borderColour.getIndex());
							}
							
						}
						
						// Apply the bottom border if set
						Border borderBottom = style.getBorder("bottom");
						if(borderBottom != null) {
							
							// Apply the style if set
							BorderStyle borderStyle = styleHelper.getBorderStyles().get(borderBottom.getType());
							if(borderStyle != null) {
								cellStyle.setBorderBottom(borderStyle);
							}
							
							// Apply the colour if set
							IndexedColors borderColour = styleHelper.getColours().get(borderBottom.getColour());
							
							// If an rgb colour has been specified then use that
							if(borderBottom.getColour().matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {
								
								String[] rgb = borderBottom.getColour().substring(4, borderBottom.getColour().length() - 1).split("\\s*,\\s*");
								cellStyle.setBottomBorderColor(new XSSFColor(new java.awt.Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])), new DefaultIndexedColorMap()));
							}
							else if(borderColour != null) {
								cellStyle.setBottomBorderColor(borderColour.getIndex());
							}
							
						}
						
						// Apply the left border if set
						Border borderLeft = style.getBorder("left");
						if(borderLeft != null) {
							
							// Apply the style if set
							BorderStyle borderStyle = styleHelper.getBorderStyles().get(borderLeft.getType());
							if(borderStyle != null) {
								cellStyle.setBorderLeft(borderStyle);
							}
							
							// Apply the colour if set
							IndexedColors borderColour = styleHelper.getColours().get(borderLeft.getColour());
							
							// If an rgb colour has been specified then use that
							if(borderLeft.getColour().matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {
								
								String[] rgb = borderLeft.getColour().substring(4, borderLeft.getColour().length() - 1).split("\\s*,\\s*");
								cellStyle.setLeftBorderColor(new XSSFColor(new java.awt.Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])), new DefaultIndexedColorMap()));
							}
							else if(borderColour != null) {
								cellStyle.setLeftBorderColor(borderColour.getIndex());
							}
							
						}
						
						// Apply the cell wrapping
						cellStyle.setWrapText(style.getWrap());
						
						// Apply the fill colour if set
						String fillColour = style.getFillColour();
						if(fillColour.length() > 0) {
							
							// If an rgb colour has been specified then use that
							if(fillColour.matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {
								
								String[] rgb = fillColour.substring(4, fillColour.length() - 1).split("\\s*,\\s*");
								cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])), new DefaultIndexedColorMap()));
							}
							else {
								cellStyle.setFillForegroundColor(styleHelper.getColours().get(fillColour).getIndex());
							}
							
						}
						
						
						// Apply the fill pattern if set
						String fillPattern = style.getFillPattern();
						if(fillPattern.length() > 0) {
							cellStyle.setFillPattern(styleHelper.getFillPatterns().get(fillPattern));
						}
						
						// Apply the font if set
						if(style.getFont() != null) {
							cellStyle.setFont(style.getFont());
						}
						
					} // End if has style
					else {
						xlCell.setCellValue(cellValue);
					}
					
					// Save the style to the cell
					xlCell.setCellStyle(cellStyle);
					
					// If a validation is set for the cell then apply the validation
					if(cell.hasAttribute("validation")) {
						
						Validation validation = validationMap.get(cell.getAttribute("validation"));
						if(validation != null) {
							CellAddress cellAddress = xlCell.getAddress();
							CellRangeAddressList rangeAddress = new CellRangeAddressList( cellAddress.getRow(), cellAddress.getRow(), cellAddress.getColumn(), cellAddress.getColumn() );
							XSSFDataValidation dvValidation = null;
							
							
							// If the type of validation is for a fixed list of values
							if(validation.getType().equals("fixed-list")) {
								
								// Get the values
								String[] values = validation.getValues();
								
								// Add the validation if the values list is not empty
								if(values.length > 0) {
									
									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(values);
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, rangeAddress);
								}
							}
							// Else if the type uses a formula to apply the list
							else if(validation.getType().equals("formula-list")) {
								
								// Get the formula
								String formula = validation.getFormula();
								
								// If a formula has been defined
								if(formula.length() > 0) {
									
									// Convert the formula to a named range
									XSSFName rangeName = xlWorkbook.createName();
									rangeName.setNameName(validation.getName());
									rangeName.setRefersToFormula(formula);
									
									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createFormulaListConstraint(validation.getName());
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, rangeAddress);
									
								}
								
							}
							
							// Apply the validation to the sheet
							dvValidation.setSuppressDropDownArrow(true);
							dvValidation.setShowErrorBox(true);
							xlSheet.addValidationData(dvValidation);
							
						}
						
					}
					
				}
				
			}
			
			// If a table has been defined then apply it to the sheet
			Element table = (Element) worksheet.getElementsByTagName("table").item(0);
			if(table != null) {
				
				// Set the area of the table using the max row and max col counts
				AreaReference tableArea = xlWorkbook.getCreationHelper().createAreaReference(
						new CellReference(0, 0), new CellReference(maxR, maxC));
				XSSFTable xlTable = xlSheet.createTable(tableArea);
				xlTable.getCTTable().addNewAutoFilter().setRef(tableArea.formatAsString());
				
				if(table.hasAttribute("name")) {
					xlTable.setName(table.getAttribute("name"));
					xlTable.setDisplayName(table.getAttribute("name"));
				}
				
				if(table.hasAttribute("style")) {
					xlTable.getCTTable().addNewTableStyleInfo();
					xlTable.getCTTable().getTableStyleInfo().setName(table.getAttribute("style"));
					XSSFTableStyleInfo tableStyle = (XSSFTableStyleInfo) xlTable.getStyle();
					tableStyle.setFirstColumn(false);
					tableStyle.setLastColumn(false);
					
					if(table.hasAttribute("colStripes")) {
						
						if(table.getAttribute("colStripes").equals("true")) {
							tableStyle.setShowColumnStripes(true);
						}
						else {
							tableStyle.setShowColumnStripes(false);
						}
						
					}
					
					if(table.hasAttribute("rowStripes")) {
						
						if(table.getAttribute("rowStripes").equals("true")) {
							tableStyle.setShowRowStripes(true);
						}
						else {
							tableStyle.setShowRowStripes(false);
						}
						
					}
				}
				
			}
			
			// Apply the auto filter if set and no table is defined
			if(worksheet.hasAttribute("autofilter")) {
				
				if(worksheet.getAttribute("autofilter").equals("true")) {
					
					if(worksheet.getElementsByTagName("table").getLength() > 0) {
						System.out.println("Skipping worksheet autofilter. Worksheet has a table with autofilters already enabled.");
					}
					else {
						xlSheet.setAutoFilter(new CellRangeAddress(0, 0, 0, maxC));
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
