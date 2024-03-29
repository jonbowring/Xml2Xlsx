/*
Copyright 2021 Jonathon Bowring

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

package com.informatica.xml2xlsx;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.FontUnderline;
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
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class AppXml2Xlsx {
	
	// Global variable declarations
	static final String outputEncoding = "UTF-8";
	static Document doc;
	static HashMap<String, XSSFCellStyle> styleMap = new HashMap<String, XSSFCellStyle>();
	static HashMap<String, StyleFormat> styleFormatMap = new HashMap<String, StyleFormat>();
	static HashMap<String, Validation> validationMap = new HashMap<String, Validation>();
	static StyleHelper styleHelper = new StyleHelper();
	static XPath xpath;
	static XSSFWorkbook xlWorkbook;
	static CreationHelper xlHelper;
	static HashMap<String, Integer> tabOrder = new HashMap<String, Integer>();
	static Boolean showProgress = false;
	
	
	public static void main(String[] args) throws Exception {

		// Declare and initialise the variables
		String src = "", tgt = "";

		// Parse the command line arguments
		for (int p = 0; p < args.length; p++) {
			
			switch (args[p]) {
			case "--src":
				src = args[p + 1];
				p++;
				break;
			case "--tgt":
				tgt = args[p + 1];
				p++;
				break;
			case "--showProgress":
				showProgress = true;
				break;
			default:
				break;
			}
			
		} // end args loop

		// Parse the source XML file
		loadXmlData(src);
		XPathFactory xpathFactory = XPathFactory.newInstance();
		xpath = xpathFactory.newXPath();

		// Initialise the target Excel workbook
		xlWorkbook = new XSSFWorkbook();
		xlHelper = xlWorkbook.getCreationHelper();

		// Get the workbook node
		System.out.println("Initialising workbook...");

		// Initialise the styles
		loadStyles();

		// Initialise the validations
		loadValidations();

		// Process the workbook data
		loadWorkbook();

		// Save and close the target Excel file
		try (OutputStream fileOut = new FileOutputStream(tgt)) {
			System.out.println("\nSaving Excel file '" + tgt + "'...");
			xlWorkbook.write(fileOut);
			xlWorkbook.close();
			System.out.println("File saved!");
		}

	} // End main

	
	// function to sort hashmap by values
	public static HashMap<String, Integer> sortByValue(HashMap<String, Integer> hm) {
		// Create a list from elements of HashMap
		List<Map.Entry<String, Integer>> list = new LinkedList<Map.Entry<String, Integer>>(hm.entrySet());

		// Sort the list
		Collections.sort(list, new Comparator<Map.Entry<String, Integer>>() {
			public int compare(Map.Entry<String, Integer> o1, Map.Entry<String, Integer> o2) {
				return (o1.getValue()).compareTo(o2.getValue());
			}
		});

		// put data from sorted list to hashmap
		HashMap<String, Integer> temp = new LinkedHashMap<String, Integer>();
		for (Map.Entry<String, Integer> aa : list) {
			temp.put(aa.getKey(), aa.getValue());
		}
		return temp;
	}
	
	
	// This function is used to read the markup from the XML file
	private static void loadXmlData(String src) throws ParserConfigurationException, SAXException, IOException {

		// Parse the source XML file
		System.out.println("Reading XML source file '" + src + "'...");
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		doc = db.parse(new File(src));
		dbf.setNamespaceAware(true);

	}
	
	
	// This function is used to load the style objects
	private static void loadStyles() throws XPathExpressionException {
		// Parse the styles
		NodeList styles = (NodeList) xpath.evaluate("/workbook/styles/style", doc, XPathConstants.NODESET);

		// If a styles array has been included
		if (styles.getLength() > 0) {

			// Loop through the styles and create the style objects
			for (int y = 0; y < styles.getLength(); y++) {

				// Initialise the style object
				Element styleEl = (Element) styles.item(y);
				String styleName = styleEl.getAttribute("name");
				XSSFCellStyle cellStyle = xlWorkbook.createCellStyle();
				XSSFDataFormat dataFmt = xlWorkbook.createDataFormat();

				// If the format is set then apply it to the cell style
				if (styleEl.getElementsByTagName("format").getLength() > 0) {
					Element format = (Element) styleEl.getElementsByTagName("format").item(0);

					if (format.hasAttribute("type")) {

						// 1 = Number, no decimal places, no thousand separator
						// 2 = Number, 2 decimal places, no thousand separator
						// 3 = Number, 0 decimal places, with thousand separator
						// 4 = Number, 2 decimal places, with thousand separator
						// 5 = Currency, 0 decimal places, locale settings
						// 7 = Currency, 2 decimal places, locale settings
						// 9 = Percentage, 0 decimal places
						// 10 = Percentage, 2 decimal places
						// 11 = Scientific, 2 decimal places
						// 12 = Fraction up to one digit (1/4)
						// 13 = Fraction up to two digits (25/26)
						// 14 = Date locale

						switch (format.getAttribute("type")) {
						case "percent":
							// Apply the pattern only if set
							if (format.hasAttribute("pattern")) {
								cellStyle.setDataFormat(
										xlHelper.createDataFormat().getFormat(format.getAttribute("pattern")));
							}
							// Else use the standard Excel format
							else {
								cellStyle.setDataFormat((short) 10);
							}
							break;
						case "currency":
							// Apply the pattern only if set
							if (format.hasAttribute("pattern")) {
								cellStyle.setDataFormat(
										xlHelper.createDataFormat().getFormat(format.getAttribute("pattern")));
							}
							// Else use the standard Excel format
							else {
								cellStyle.setDataFormat((short) 7);
							}
							break;
						case "scientific":
							cellStyle.setDataFormat((short) 11);
							break;
						case "fraction":
							cellStyle.setDataFormat((short) 12);
							break;
						case "formula":
							// Do nothing
							break;
						case "string":
							cellStyle.setDataFormat(dataFmt.getFormat("@"));
							break;
						case "int":
							// If the thousands separator flag is set then use that
							if (format.hasAttribute("separator")) {
								if (format.getAttribute("separator").equals("true")) {
									cellStyle.setDataFormat((short) 3);
								}
							} else {
								cellStyle.setDataFormat((short) 1);
							}
							break;
						case "float":
							// Apply the pattern only if set
							if (format.hasAttribute("pattern")) {
								cellStyle.setDataFormat(
										xlHelper.createDataFormat().getFormat(format.getAttribute("pattern")));
							}
							// Else use the standard Excel format
							else {
								// If the thousands separator flag is set then use that
								if (format.hasAttribute("separator")) {
									if (format.getAttribute("separator").equals("true")) {
										cellStyle.setDataFormat((short) 4);
									}
								} else {
									cellStyle.setDataFormat((short) 2);
								}
							}
							break;
						case "date":
							// If a custom date pattern is set then apply that
							if (format.hasAttribute("pattern")) {
								cellStyle.setDataFormat(
										xlHelper.createDataFormat().getFormat(format.getAttribute("pattern")));
							}
							// Else use the standard Excel date locale format
							else {
								cellStyle.setDataFormat((short) 14);
							}
							break;

						case "datetime":
							if (format.hasAttribute("pattern")) {
								cellStyle.setDataFormat(
										xlHelper.createDataFormat().getFormat(format.getAttribute("pattern")));
							}
							// Else use the standard Excel date locale format
							else {
								cellStyle.setDataFormat((short) 14);
							}
							break;

						default:
							// Do nothing
							break;
						} // End data type switch

					}
				}

				// If the style has an attribute element
				if (styleEl.getElementsByTagName("align").getLength() > 0) {

					Element align = (Element) styleEl.getElementsByTagName("align").item(0);

					// If it has the valign attribute then save it
					if (align.hasAttribute("vertical")) {

						switch (align.getAttribute("vertical")) {
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

						} // End valign switch

					}

					// If it has the halign attribute then save it
					if (align.hasAttribute("horizontal")) {

						switch (align.getAttribute("horizontal")) {
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

						} // End halign switch

					}

				}

				// If the style has a fill element
				if (styleEl.getElementsByTagName("fill").getLength() > 0) {

					Element fill = (Element) styleEl.getElementsByTagName("fill").item(0);

					// If it has the colour element then save it
					if (fill.hasAttribute("colour")) {
						String fillColour = fill.getAttribute("colour");

						// If an rgb colour has been specified then use that
						if (fillColour.matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {

							String[] rgb = fillColour.substring(4, fillColour.length() - 1).split("\\s*,\\s*");
							cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(Integer.parseInt(rgb[0]),
									Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])), new DefaultIndexedColorMap()));
						} else {
							cellStyle.setFillForegroundColor(styleHelper.getColours().get(fillColour).getIndex());
						}

					}

					// If it has the colour and pattern defined then save it
					if (fill.hasAttribute("colour") && fill.hasAttribute("pattern")) {

						// Apply the fill pattern if set
						String fillPattern = fill.getAttribute("pattern");
						if (fillPattern.length() > 0) {
							cellStyle.setFillPattern(styleHelper.getFillPatterns().get(fillPattern));
						}

					}
					// Else if it has the the colour but no pattern then use a default pattern
					else if (fill.hasAttribute("colour") && !fill.hasAttribute("pattern")) {
						cellStyle.setFillPattern(styleHelper.getFillPatterns().get("solid-foreground"));

					}

				} // End if style has fill element

				// If the style has a wrap element then set wrap to true
				if (styleEl.getElementsByTagName("wrap").getLength() > 0) {

					cellStyle.setWrapText(true);

				}

				// If the style has border elements
				NodeList borders = styleEl.getElementsByTagName("border");
				if (borders.getLength() > 0) {

					// Loop through all of the borders
					for (int b = 0; b < borders.getLength(); b++) {

						// Initialise the border object
						Element borderEl = (Element) borders.item(b);
						String borderPos = "", borderType = "", borderColour = "";
						BorderStyle borderStyle = null;
						IndexedColors idxBorderColour = null;

						if (borderEl.hasAttribute("pos")) {
							borderPos = borderEl.getAttribute("pos");
						}

						if (borderEl.hasAttribute("type")) {
							borderType = borderEl.getAttribute("type");
						}

						if (borderEl.hasAttribute("colour")) {
							borderColour = borderEl.getAttribute("colour");
						}

						switch (borderPos) {
						case "top":
							// Apply the style if set
							borderStyle = styleHelper.getBorderStyles().get(borderType);
							if (borderStyle != null) {
								cellStyle.setBorderTop(borderStyle);
							}

							// Apply the colour if set
							idxBorderColour = styleHelper.getColours().get(borderColour);

							// If an rgb colour has been specified then use that
							if (borderColour.matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {

								String[] rgb = borderColour.substring(4, borderColour.length() - 1).split("\\s*,\\s*");
								cellStyle
										.setTopBorderColor(
												new XSSFColor(
														new java.awt.Color(Integer.parseInt(rgb[0]),
																Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])),
														new DefaultIndexedColorMap()));
							} else if (idxBorderColour != null) {
								cellStyle.setTopBorderColor(idxBorderColour.getIndex());
							}

							break;

						case "right":
							// Apply the style if set
							borderStyle = styleHelper.getBorderStyles().get(borderType);
							if (borderStyle != null) {
								cellStyle.setBorderRight(borderStyle);
							}

							// Apply the colour if set
							idxBorderColour = styleHelper.getColours().get(borderColour);

							// If an rgb colour has been specified then use that
							if (borderColour.matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {

								String[] rgb = borderColour.substring(4, borderColour.length() - 1).split("\\s*,\\s*");
								cellStyle
										.setRightBorderColor(
												new XSSFColor(
														new java.awt.Color(Integer.parseInt(rgb[0]),
																Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])),
														new DefaultIndexedColorMap()));
							} else if (idxBorderColour != null) {
								cellStyle.setRightBorderColor(idxBorderColour.getIndex());
							}

							break;

						case "bottom":
							// Apply the style if set
							borderStyle = styleHelper.getBorderStyles().get(borderType);
							if (borderStyle != null) {
								cellStyle.setBorderBottom(borderStyle);
							}

							// Apply the colour if set
							idxBorderColour = styleHelper.getColours().get(borderColour);

							// If an rgb colour has been specified then use that
							if (borderColour.matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {

								String[] rgb = borderColour.substring(4, borderColour.length() - 1).split("\\s*,\\s*");
								cellStyle
										.setBottomBorderColor(
												new XSSFColor(
														new java.awt.Color(Integer.parseInt(rgb[0]),
																Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])),
														new DefaultIndexedColorMap()));
							} else if (idxBorderColour != null) {
								cellStyle.setBottomBorderColor(idxBorderColour.getIndex());
							}

							break;

						case "left":
							// Apply the style if set
							borderStyle = styleHelper.getBorderStyles().get(borderType);
							if (borderStyle != null) {
								cellStyle.setBorderLeft(borderStyle);
							}

							// Apply the colour if set
							idxBorderColour = styleHelper.getColours().get(borderColour);

							// If an rgb colour has been specified then use that
							if (borderColour.matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {

								String[] rgb = borderColour.substring(4, borderColour.length() - 1).split("\\s*,\\s*");
								cellStyle
										.setLeftBorderColor(
												new XSSFColor(
														new java.awt.Color(Integer.parseInt(rgb[0]),
																Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])),
														new DefaultIndexedColorMap()));
							} else if (idxBorderColour != null) {
								cellStyle.setLeftBorderColor(idxBorderColour.getIndex());
							}

							break;

						default:
							break;
						}

					} // End of borders loop

				} // End if borders length > 0

				// Add the font styles if set
				Element fontEl = (Element) styleEl.getElementsByTagName("font").item(0);
				if (fontEl != null) {

					// Initialise the font
					XSSFFont font = xlWorkbook.createFont();

					// Get the font settings
					String fontName = fontEl.getAttribute("name");
					String fontSize = fontEl.getAttribute("size");
					String fontColour = fontEl.getAttribute("colour");
					Element fontItalic = (Element) fontEl.getElementsByTagName("italic").item(0);
					Element fontStrike = (Element) fontEl.getElementsByTagName("strikeout").item(0);
					Element fontBold = (Element) fontEl.getElementsByTagName("bold").item(0);
					Element fontUnderline = (Element) fontEl.getElementsByTagName("underline").item(0);

					// Set the font name if set
					if (fontName.length() > 0) {
						font.setFontName(fontName);
					}

					// Set the font size if set
					if (fontSize.length() > 0) {
						font.setFontHeightInPoints(Short.parseShort(fontSize));
					}

					// Set the font colour if set
					if (fontColour.length() > 0) {

						// If an rgb colour has been specified then use that
						if (fontColour.matches("^rgb\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*\\)$")) {
							String[] rgb = fontColour.substring(4, fontColour.length() - 1).split("\\s*,\\s*");
							font.setColor(new XSSFColor(new java.awt.Color(Integer.parseInt(rgb[0]),
									Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2])), new DefaultIndexedColorMap()));
						}

						// Else lookup the colour index
						else {
							font.setColor(styleHelper.getColours().get(fontColour).getIndex());
						}

					}

					// Set the font italic if set
					if (fontItalic != null) {
						font.setItalic(true);
					}

					// Set the font strikeout if set
					if (fontStrike != null) {
						font.setStrikeout(true);
					}

					// Set the font bold if set
					if (fontBold != null) {
						font.setBold(true);
					}

					// Set the font underline if set
					if (fontUnderline != null) {
						font.setUnderline(FontUnderline.SINGLE);
					}

					// Save the font
					cellStyle.setFont(font);

				} // End if has font element

				// Add the format styles if set
				Element formatEl = (Element) styleEl.getElementsByTagName("format").item(0);
				if (formatEl != null) {
					if (formatEl.hasAttribute("type")) {
						if (formatEl.hasAttribute("formula")) {
							styleFormatMap.put(styleName, new StyleFormat(formatEl.getAttribute("type"),
									Boolean.parseBoolean(formatEl.getAttribute("formula"))));
						} else {
							styleFormatMap.put(styleName, new StyleFormat(formatEl.getAttribute("type")));
						}
					}
				}

				// Add the current style to the hash map
				styleMap.put(styleName, cellStyle);

			} // End of styles for loop

		} // End of if have styles check
	}
	
	
	// This function is used to load the validation objects
	private static void loadValidations() throws XPathExpressionException {
		
		// Parse the validations
		NodeList validations = (NodeList) xpath.evaluate("/workbook/validations/validation", doc,
				XPathConstants.NODESET);

		// If a validations array has been included
		if (validations.getLength() > 0) {

			// Loop through the validations and create the validation objects
			for (int v = 0; v < validations.getLength(); v++) {

				// Initialise the validation object
				Element validationEl = (Element) validations.item(v);
				String validationType = validationEl.getElementsByTagName("type").item(0).getTextContent();
				Validation validation = new Validation(validationEl.getAttribute("name"), validationType);

				// If set get and set the formula
				Element formulaEl = (Element) validationEl.getElementsByTagName("formula").item(0);
				if (formulaEl != null) {
					validation.setFormula(formulaEl.getTextContent());
				}

				// If it is a length or value validation type
				if (validationType.equals("length") || validationType.equals("numerical") || validationType.equals("date")) {

					// If set get the operator and add it to the validation
					Element operatorEl = (Element) validationEl.getElementsByTagName("operator").item(0);
					if (operatorEl != null) {

						switch (operatorEl.getTextContent()) {
						case "EQUAL":
							validation.setOperator(OperatorType.EQUAL);
							break;
						case "GREATER_OR_EQUAL":
							validation.setOperator(OperatorType.GREATER_OR_EQUAL);
							break;
						case "LESS_OR_EQUAL":
							validation.setOperator(OperatorType.LESS_OR_EQUAL);
							break;
						case "GREATER_THAN":
							validation.setOperator(OperatorType.GREATER_THAN);
							break;
						case "LESS_THAN":
							validation.setOperator(OperatorType.LESS_THAN);
							break;
						case "NOT_EQUAL":
							validation.setOperator(OperatorType.NOT_EQUAL);
							break;
						case "BETWEEN":
							validation.setOperator(OperatorType.BETWEEN);
							break;
						case "NOT_BETWEEN":
							validation.setOperator(OperatorType.NOT_BETWEEN);
							break;
						default:
							// do something
							break;
						} // End of operator switch statement

					} // End if operator element is not null

					// If it is a length validation type
					if (validationType.equals("length")) {

						// If set get and set the length value
						Element lengthValueEl = (Element) validationEl.getElementsByTagName("value").item(0);
						if (lengthValueEl != null) {
							validation.setLengthValue(Integer.parseInt(lengthValueEl.getTextContent()));
						}

						// If set get and set the length min value
						Element lengthMinValueEl = (Element) validationEl.getElementsByTagName("min").item(0);
						if (lengthMinValueEl != null) {
							validation.setLengthMin(Integer.parseInt(lengthMinValueEl.getTextContent()));
						}

						// If set get and set the length max value
						Element lengthMaxValueEl = (Element) validationEl.getElementsByTagName("max").item(0);
						if (lengthMaxValueEl != null) {
							validation.setLengthMax(Integer.parseInt(lengthMaxValueEl.getTextContent()));
						}
					} // End if length validation

					// If a numerical validation type
					else if (validationType.equals("numerical")) {

						// If set get and set the exact value
						Element numExactEl = (Element) validationEl.getElementsByTagName("value").item(0);
						if (numExactEl != null) {
							validation.setNumExact(Float.parseFloat(numExactEl.getTextContent()));
						}

						// If set get and set the min value
						Element numMinEl = (Element) validationEl.getElementsByTagName("min").item(0);
						if (numMinEl != null) {
							validation.setNumMin(Float.parseFloat(numMinEl.getTextContent()));
						}

						// If set get and set the max value
						Element numMaxEl = (Element) validationEl.getElementsByTagName("max").item(0);
						if (numMaxEl != null) {
							validation.setNumMax(Float.parseFloat(numMaxEl.getTextContent()));
						}
					} // End if numerical validation
					
					// TODO
					// If a date validation type
					else if (validationType.equals("date")) {

						// If set get and set the exact value
						Element dateExactEl = (Element) validationEl.getElementsByTagName("value").item(0);
						if (dateExactEl != null) {
							validation.setDateExact(dateExactEl.getTextContent());
						}

						// If set get and set the min value
						Element dateMinEl = (Element) validationEl.getElementsByTagName("min").item(0);
						if (dateMinEl != null) {
							validation.setDateMin(dateMinEl.getTextContent());
						}

						// If set get and set the max value
						Element dateMaxEl = (Element) validationEl.getElementsByTagName("max").item(0);
						if (dateMaxEl != null) {
							validation.setDateMax(dateMaxEl.getTextContent());
						}
					} // End if date validation

				} // End if length or value validation

				// If set get the list of validation values
				Element valuesArrEl = (Element) validationEl.getElementsByTagName("values").item(0);
				if (valuesArrEl != null) {

					NodeList valuesEl = valuesArrEl.getElementsByTagName("value");

					if (valuesEl.getLength() > 0) {
						List<String> valuesList = new ArrayList<String>();
						for (int a = 0; a < valuesEl.getLength(); a++) {

							Element valueEl = (Element) valuesEl.item(a);
							valuesList.add(valueEl.getTextContent());

						}

						// Save the list of values to the validation
						validation.setValues(valuesList.toArray(new String[valuesEl.getLength()]));

					} // End if values length > 0

				} // End if values array is not null

				// Save the validation to the map
				validationMap.put(validation.getName(), validation);

			} // End validations loop

		} // End if validations length > 0
	}
	
	
	// This function is used to create the output workbook
	private static void loadWorkbook() throws XPathExpressionException, ParseException {
		
		// Get all worksheets in the workbook and loop through them
		NodeList worksheets = (NodeList) xpath.evaluate("/workbook/worksheet", doc, XPathConstants.NODESET);
		for (int s = 0; s < worksheets.getLength(); s++) {

			// Get the current worksheet
			Element worksheet = (Element) worksheets.item(s);
			String sheetName = worksheet.getAttribute("name");
			System.out.println("Adding worksheet '" + sheetName + "'...");

			// Initialise the target Excel worksheet and data validation helper
			XSSFSheet xlSheet = xlWorkbook.createSheet(sheetName);
			XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(xlSheet);

			// If the worksheet order is set then update the workbook order
			if (worksheet.hasAttribute("order")) {
				Integer order = Integer.parseInt(worksheet.getAttribute("order"));
				tabOrder.put(sheetName, order);
			} else {
				tabOrder.put(sheetName, -1);
			}

			// If the gridlines flag is set then add it to the worksheet
			if (worksheet.hasAttribute("gridlines")) {
				Boolean gridlines = Boolean.parseBoolean(worksheet.getAttribute("gridlines"));
				if (gridlines) {
					xlSheet.setDisplayGridlines(true);
				} else {
					xlSheet.setDisplayGridlines(false);
				}
			}

			// Get all rows in the current worksheet and loop through them
			NodeList rows = worksheet.getElementsByTagName("row");
			int maxR = 0, maxC = 0;
			Double complete = 0.0, total = (double) rows.getLength();

			// Get any column settings for the current worksheet
			NodeList cols = worksheet.getElementsByTagName("column");
			ArrayList<Integer> ignoreAutoFit = new ArrayList<Integer>();
			for (int w = 0; w < cols.getLength(); w++) {

				// Get the current column
				Element col = (Element) cols.item(w);

				// If the width attribute is set then set the width and add it to the map to be
				// ignored by auto fit
				if (col.hasAttribute("index") && col.hasAttribute("width")) {
					int colIndex = Integer.parseInt(col.getAttribute("index"));
					int colWidth = Integer.parseInt(col.getAttribute("width"));
					xlSheet.setColumnWidth(colIndex, colWidth);
					ignoreAutoFit.add(colIndex);
				}

				// If the column has a default style then apply it
				if (col.hasAttribute("index") && col.hasAttribute("style")) {
					int colIndex = Integer.parseInt(col.getAttribute("index"));
					xlSheet.getColumnHelper().setColDefaultStyle(colIndex, styleMap.get(col.getAttribute("style")));
				}
			}

			// If a table has been defined then initialise it
			Element table = (Element) worksheet.getElementsByTagName("table").item(0);
			XSSFTable xlTable = null;
			if (table != null) {

				// Set the area of the table using the max row and max col counts
				AreaReference tableArea = xlWorkbook.getCreationHelper().createAreaReference(new CellReference(0, 0),
						new CellReference(1, 1));

				xlTable = xlSheet.createTable(tableArea);
				xlTable.getCTTable().addNewAutoFilter().setRef(tableArea.formatAsString());

				if (table.hasAttribute("name")) {
					xlTable.setName(table.getAttribute("name"));
					xlTable.setDisplayName(table.getAttribute("name"));
				}

				if (table.hasAttribute("style")) {
					xlTable.getCTTable().addNewTableStyleInfo();
					xlTable.getCTTable().getTableStyleInfo().setName(table.getAttribute("style"));
					XSSFTableStyleInfo tableStyle = (XSSFTableStyleInfo) xlTable.getStyle();
					tableStyle.setFirstColumn(false);
					tableStyle.setLastColumn(false);

					if (table.hasAttribute("colStripes")) {

						if (table.getAttribute("colStripes").equals("true")) {
							tableStyle.setShowColumnStripes(true);
						} else {
							tableStyle.setShowColumnStripes(false);
						}

					}

					if (table.hasAttribute("rowStripes")) {

						if (table.getAttribute("rowStripes").equals("true")) {
							tableStyle.setShowRowStripes(true);
						} else {
							tableStyle.setShowRowStripes(false);
						}

					}

				} // End if table has style

			} // End if has table

			for (int r = 0; r < rows.getLength(); r++) {

				// display the progress bar if the argument is passed
				if (showProgress) {
					// Calculate the progress percentage
					Double progress = (double) r + 1;
					complete = (progress / total) * 100;
					System.out.print("[");
					for (int n = 0; n < 100; n++) {
						if (n <= complete) {
							System.out.print("=");
						} else {
							System.out.print(" ");
						}
					}
					System.out.print("] " + Math.round(complete) + "% (" + (r + 1) + "/" + rows.getLength() + ")\r");
				}

				// Update the max row count
				if (r > maxR) {
					maxR = r;
				}

				// Get the current row
				Element row = (Element) rows.item(r);

				// Initialise the target row
				XSSFRow xlRow = xlSheet.createRow(r);

				// Get all cells in the current row and loop through them
				NodeList cells = row.getElementsByTagName("cell");
				for (int c = 0; c < cells.getLength(); c++) {

					// Update the max col count
					if (c > maxC) {
						maxC = c;
					}

					// Get the current cell
					Element cell = (Element) cells.item(c);
					String cellValue = cell.getTextContent();

					// Initialise the target Excel cell and add the value
					XSSFCell xlCell = xlRow.createCell(c);

					// Add the column to the table if a table is in use
					if (xlTable != null && r == 0) {
						xlTable.createColumn(cellValue, c);
					}

					// If a cell specifies a default style then apply it to the entire column
					if (cell.hasAttribute("columnStyle")) {
						xlSheet.getColumnHelper().setColDefaultStyle(c, styleMap.get(cell.getAttribute("columnStyle")));
					}

					// If a cell style has been applied then add it to the cell
					if (cell.hasAttribute("style")) {

						// Apply the cell format if set
						String styleFormat = null;
						if (styleFormatMap.containsKey(cell.getAttribute("style"))) {
							styleFormat = styleFormatMap.get(cell.getAttribute("style")).getType();
						}

						if (styleFormat != null) {

							if (styleFormat.length() > 0) {

								// If the style is a formula then set the value as a formula
								Boolean styleIsFormula = styleFormatMap.get(cell.getAttribute("style")).getIsFormula();
								if (styleIsFormula) {
									if (cellValue == null || cellValue.length() == 0) {
										xlCell.setCellValue("");
									} else {
										xlCell.setCellFormula(cellValue);
									}
								}
								// Else set the value by the format type
								else {
									switch (styleFormat) {
									case "currency":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											Double cellDouble = Double.parseDouble(cellValue);
											xlCell.setCellValue(cellDouble);
										}
										break;
									case "scientific":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											Double cellDouble = Double.parseDouble(cellValue);
											xlCell.setCellValue(cellDouble);
										}
										break;
									case "fraction":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											Double cellDouble = Double.parseDouble(cellValue);
											xlCell.setCellValue(cellDouble);
										}
										break;
									case "percent":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											Double cellDouble = Double.parseDouble(cellValue);
											xlCell.setCellValue(cellDouble);
										}
										break;
									case "string":
										xlCell.setCellValue(cellValue);
										break;
									case "int":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											int cellInt = Integer.parseInt(cellValue);
											xlCell.setCellValue(cellInt);
										}
										break;
									case "float":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											Double cellDouble = Double.parseDouble(cellValue);
											xlCell.setCellValue(cellDouble);
										}
										break;
									case "date":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
											Date cellDate = fmt.parse(cellValue);
											xlCell.setCellValue(cellDate);

										}

										break;

									case "datetime":
										if (cellValue == null || cellValue.length() == 0) {
											xlCell.setCellValue("");
										} else {
											SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
											Date cellDate = fmt.parse(cellValue);
											xlCell.setCellValue(cellDate);

										}

										break;

									default:
										xlCell.setCellValue(cellValue);
										break;
									} // End data type switch
								} // End else is not formula

							}

						} // End if cell has format
						else {
							xlCell.setCellValue(cellValue);
						}

						// Save the style to the cell
						xlCell.setCellStyle(styleMap.get(cell.getAttribute("style")));

					} // End if has style
					else {
						xlCell.setCellValue(cellValue);

					} // End else has style

					/*
					 * Manage the cell validations
					 */

					// If a validation is set for the cell then apply the validation
					if (cell.hasAttribute("validation")) {

						Validation validation = validationMap.get(cell.getAttribute("validation"));
						if (validation != null) {
							CellAddress cellAddress = xlCell.getAddress();
							CellRangeAddressList rangeAddress = new CellRangeAddressList(cellAddress.getRow(),
									cellAddress.getRow(), cellAddress.getColumn(), cellAddress.getColumn());
							XSSFDataValidation dvValidation = null;

							// If the type of validation is for a fixed list of values
							if (validation.getType().equals("fixed-list")) {

								// Get the values
								String[] values = validation.getValues();

								// Add the validation if the values list is not empty
								if (values.length > 0) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createExplicitListConstraint(values);
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);
								}

							} // End if fixed-list

							// Else if the type uses a formula to apply the list
							else if (validation.getType().equals("formula-list")) {

								// Get the formula
								String formula = validation.getFormula();

								// If a formula has been defined
								if (formula.length() > 0) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createFormulaListConstraint(formula);
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);

								}

							} // End if formula list

							// Else if the type uses a length validation
							else if (validation.getType().equals("length")) {

								// Get the operator and length values
								int operator = validation.getOperator();
								int lengthValue = validation.getLengthValue();
								int lengthMin = validation.getLengthMin();
								int lengthMax = validation.getLengthMax();

								// If the operation only requires a single value
								if ((operator == OperatorType.EQUAL || operator == OperatorType.GREATER_OR_EQUAL
										|| operator == OperatorType.GREATER_THAN
										|| operator == OperatorType.LESS_OR_EQUAL || operator == OperatorType.LESS_THAN
										|| operator == OperatorType.NOT_EQUAL) && lengthValue > -1) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createTextLengthConstraint(operator, Integer.toString(lengthValue), null);
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);

								}
								// If the operation requires two values
								else if ((operator == OperatorType.BETWEEN || operator == OperatorType.NOT_BETWEEN)
										&& lengthMin > -1 && lengthMax > -1) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createTextLengthConstraint(operator, Integer.toString(lengthMin),
													Integer.toString(lengthMax));
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);

								}

							} // End if length validation

							// Else if the type uses a numerical validation
							else if (validation.getType().equals("numerical")) {

								// Get the operator and length values
								int operator = validation.getOperator();
								Float valueExact = validation.getNumExact();
								Float valueMin = validation.getNumMin();
								Float valueMax = validation.getNumMax();

								// If the operation only requires a single value
								if ((operator == OperatorType.EQUAL || operator == OperatorType.GREATER_OR_EQUAL
										|| operator == OperatorType.GREATER_THAN
										|| operator == OperatorType.LESS_OR_EQUAL || operator == OperatorType.LESS_THAN
										|| operator == OperatorType.NOT_EQUAL) && valueExact != null) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createNumericConstraint(
													XSSFDataValidationConstraint.ValidationType.DECIMAL, operator,
													Float.toString(valueExact), null);
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);

								}
								// If the operation requires two values
								else if ((operator == OperatorType.BETWEEN || operator == OperatorType.NOT_BETWEEN)
										&& valueMin != null && valueMax != null) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createNumericConstraint(
													XSSFDataValidationConstraint.ValidationType.DECIMAL, operator,
													Float.toString(valueMin), Float.toString(valueMax));
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);

								}

							} // End if numerical validation
							
							// TODO
							// Else if the type uses a date validation
							else if (validation.getType().equals("date")) {

								// Get the operator and length values
								int operator = validation.getOperator();
								String valueExact = validation.getDateExactFunc();
								String valueMin = validation.getDateMinFunc();
								String valueMax = validation.getDateMaxFunc();

								// If the operation only requires a single value
								if ((operator == OperatorType.EQUAL || operator == OperatorType.GREATER_OR_EQUAL
										|| operator == OperatorType.GREATER_THAN
										|| operator == OperatorType.LESS_OR_EQUAL || operator == OperatorType.LESS_THAN
										|| operator == OperatorType.NOT_EQUAL) && valueExact != null) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createDateConstraint(operator, valueExact, null, "yyyy-MM-dd");
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);

								}
								// If the operation requires two values
								else if ((operator == OperatorType.BETWEEN || operator == OperatorType.NOT_BETWEEN)
										&& valueMin != null && valueMax != null) {

									// Build the validation
									XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
											.createDateConstraint(operator, valueMin, valueMax, "yyyy-MM-dd");
									dvValidation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint,
											rangeAddress);

								}

							} // End if date validation

							// Apply the validation to the sheet
							dvValidation.setSuppressDropDownArrow(true);
							dvValidation.setShowErrorBox(true);
							xlSheet.addValidationData(dvValidation);

						} // End if validation is not null

					} // End if cell has validation

				} // End of cells loop

			} // End of rows loop

			/*
			 * Manage the worksheet table
			 */

			// If a table has been defined then update the range
			if (table != null) {

				// Set the area of the table using the max row and max col counts
				AreaReference tableArea = null;

				// If there are rows available then set the area using the row count
				if (maxR > 0) {
					tableArea = xlWorkbook.getCreationHelper().createAreaReference(new CellReference(0, 0),
							new CellReference(maxR, maxC));
				}
				// Else set the area with an empty row
				else {
					tableArea = xlWorkbook.getCreationHelper().createAreaReference(new CellReference(0, 0),
							new CellReference(1, maxC));
				}

				xlTable.setArea(tableArea);

			} // End if has table

			/*
			 * Manage the worksheet autofilters
			 */

			// Apply the auto filter if set and no table is defined
			if (worksheet.hasAttribute("autofilter")) {

				if (worksheet.getAttribute("autofilter").equals("true")) {

					if (worksheet.getElementsByTagName("table").getLength() > 0) {
						System.out.println(
								"\nSkipping worksheet autofilter. Worksheet has a table with autofilters already enabled.");
					} else {
						xlSheet.setAutoFilter(new CellRangeAddress(0, 0, 0, maxC));
					}

				}

			} // End if has autofilter

			/*
			 * Manage the worksheet column autofit
			 */

			// Apply the column auto fit if set
			if (worksheet.hasAttribute("autofit")) {

				if (worksheet.getAttribute("autofit").equals("true")) {

					for (int i = 0; i <= maxC; i++) {

						// Auto fit the column as long as it is not in the ignore auto fit list
						if (!ignoreAutoFit.contains(i)) {
							xlSheet.autoSizeColumn(i);
						}

					}

				}

			} // End if has autofit

			/*
			 * Manage the worksheet visibility
			 */

			// Apply the worksheet visibility if set
			if (worksheet.hasAttribute("hidden")) {

				if (worksheet.getAttribute("hidden").equals("true")) {

					xlWorkbook.setSheetHidden(xlWorkbook.getSheetIndex(sheetName), true);

				}

			} // End if has autofit

			/*
			 * Manage the pivot tables
			 */

			// Get all pivots in the workbook and loop trhough them
			NodeList pivots = worksheet.getElementsByTagName("pivot");
			for (int p = 0; p < pivots.getLength(); p++) {

				// Get the current pivot table
				Element pivot = (Element) pivots.item(p);

				// Get the pivot table properties
				String location = pivot.getAttribute("location");
				String dataArea = pivot.getAttribute("dataArea");
				String dataSheet = pivot.getAttribute("dataSheet");

				// Initialise the pivot table
				XSSFPivotTable pivotTable = null;
				if (pivot.hasAttribute("dataTable")) {
					pivotTable = xlSheet.createPivotTable(
							xlWorkbook.getTable(pivot.getAttribute("dataTable")).getArea(), new CellReference(location),
							xlWorkbook.getSheet(dataSheet));
				} else {
					pivotTable = xlSheet.createPivotTable(new AreaReference(dataArea, SpreadsheetVersion.EXCEL2007),
							new CellReference(location), xlWorkbook.getSheet(dataSheet));
				}

				// Add the columns to be used for grouping
				Element groupby = (Element) pivot.getElementsByTagName("groupby").item(0);
				NodeList groupbyCols = groupby.getElementsByTagName("column");
				for (int gc = 0; gc < groupbyCols.getLength(); gc++) {

					// Get the column number
					Element groupbyCol = (Element) groupbyCols.item(gc);
					int colIdx = Integer.parseInt(groupbyCol.getAttribute("index"));

					// Add the column to the row labels
					pivotTable.addRowLabel(colIdx);

				} // End of grouby loop

				// Add the columns to be used for calculation
				Element agg = (Element) pivot.getElementsByTagName("aggregate").item(0);
				NodeList aggCols = agg.getElementsByTagName("column");
				for (int ac = 0; ac < aggCols.getLength(); ac++) {

					// Get the column number
					Element aggCol = (Element) aggCols.item(ac);
					int colIdx = Integer.parseInt(aggCol.getAttribute("index"));
					String colAction = aggCol.getAttribute("action");

					// Add the column to the calculation
					switch (colAction) {
					case "AVERAGE":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, colIdx);
						}
						break;
					case "COUNT":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, colIdx);
						}
						break;
					case "COUNT_NUMS":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.COUNT_NUMS, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.COUNT_NUMS, colIdx);
						}
						break;
					case "MAX":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.MAX, colIdx, aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.MAX, colIdx);
						}
						break;
					case "MIN":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.MIN, colIdx, aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.MIN, colIdx);
						}
						break;
					case "PRODUCT":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.PRODUCT, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.PRODUCT, colIdx);
						}
						break;
					case "STD_DEV":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.STD_DEV, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.STD_DEV, colIdx);
						}
						break;
					case "STD_DEVP":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.STD_DEVP, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.STD_DEVP, colIdx);
						}
						break;
					case "SUM":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.SUM, colIdx, aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.SUM, colIdx);
						}
						break;
					case "VAR":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.VAR, colIdx, aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.VAR, colIdx);
						}
						break;
					case "VARP":
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.VARP, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.VARP, colIdx);
						}
						break;
					default:
						if (aggCol.hasAttribute("name")) {
							pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, colIdx,
									aggCol.getAttribute("name"));
						} else {
							pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, colIdx);
						}
						break;
					}

				} // End of aggregate loop

				// Add the columns to be used for grouping
				Element filter = (Element) pivot.getElementsByTagName("filter").item(0);
				NodeList filterCols = filter.getElementsByTagName("column");
				for (int fc = 0; fc < filterCols.getLength(); fc++) {

					// Get the column number
					Element filterCol = (Element) filterCols.item(fc);
					int colIdx = Integer.parseInt(filterCol.getAttribute("index"));

					// Add the column to the filter
					pivotTable.addReportFilter(colIdx);

				} // End of filter loop

			} // End of pivots loop

		} // End of worksheets loop

		/*
		 * Manage the tab order
		 */

		// Sort the tab positions
		if (tabOrder.size() > 0) {
			tabOrder = sortByValue(tabOrder);
			int t = 0;
			for (Map.Entry<String, Integer> entry : tabOrder.entrySet()) {

				// Remove any default tab selection
				xlWorkbook.getSheet(entry.getKey()).setSelected(false);

				if (entry.getValue() >= 0) {

					// Set the tab order
					xlWorkbook.setSheetOrder(entry.getKey(), entry.getValue());

					// Select the first tab
					if (t == 0) {
						xlWorkbook.setSelectedTab(xlWorkbook.getSheetIndex(entry.getKey()));
						xlWorkbook.setActiveSheet(xlWorkbook.getSheetIndex(entry.getKey()));
					}

					t++;
				}
			}
		}
	}

}
