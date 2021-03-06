package com.informatica.xml2xlsx;

import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Style {
	
	/*
	 * Attribute declaration
	 */
	
	private String name, valign, halign, formatType, formatPattern, fillColour, fillPattern;
	private Boolean wrap;
	private HashMap<String, Border> borderMap;
	private Font font;
	
	/*
	 * Constructors
	 */
	
	public Style(String name, XSSFWorkbook workbook) {
		this.name = name;
		this.valign = "";
		this.halign = "";
		this.formatType = "";
		this.formatPattern = "";
		this.fillColour = "";
		this.fillPattern = "";
		this.wrap = false;
		this.borderMap = new HashMap<String, Border>();
		this.font = workbook.createFont();
	}
	
	/*
	 * Getters
	 */
	
	public String getName() {
		return this.name;
	}
	
	public String getVAlign() {
		return this.valign;
	}
	
	public String getHAlign() {
		return this.halign;
	}
	
	public String getFormatType() {
		return this.formatType;
	}
	
	public String getFormatPattern() {
		return this.formatPattern;
	}
	
	public String getFillColour() {
		return this.fillColour;
	}
	
	public String getFillPattern() {
		return this.fillPattern;
	}
	
	public Boolean getWrap() {
		return this.wrap;
	}
	
	public HashMap<String, Border> getBorderMap() {
		return this.borderMap;
	}
	
	public Font getFont() {
		return this.font;
	}
	
	/*
	 * Setters
	 */
	
	public void setName(String name) {
		this.name = name;
	}
	
	public void setVAlign(String valign) {
		this.valign = valign;
	}
	
	public void setHAlign(String halign) {
		this.halign = halign;
	}
	
	public void setFormatType(String formatType) {
		this.formatType = formatType;
	}
	
	public void setFormatPattern(String formatPattern) {
		this.formatPattern = formatPattern;
	}
	
	public void setFillColour(String fillColour) {
		this.fillColour = fillColour;
	}
	
	public void setFillPattern(String fillPattern) {
		this.fillPattern = fillPattern;
	}
	
	public void setWrap(Boolean wrap) {
		this.wrap = wrap;
	}
	
	public void setBorderMap(HashMap<String, Border> borderMap) {
		this.borderMap = borderMap;
	}
	
	public void setFont(Font font) {
		this.font.setItalic(false);
		this.font.setStrikeout(false);
		this.font = font;
	}
	
	/*
	 * Border functions
	 */
	
	public Border getBorder(String pos) {
		return this.borderMap.get(pos);
	}
	
	public void addBorder(Border border) {
		this.borderMap.put(border.getPos(), border);
	}
	
	public void removeBorder(String pos) {
		this.borderMap.remove(pos);
	}
	
}
