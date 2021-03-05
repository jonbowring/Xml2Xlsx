package com.informatica.xml2xlsx;

import java.util.HashMap;

public class Style {
	
	/*
	 * Attribute declaration
	 */
	
	private String name, valign, halign, format, pattern;
	private Boolean wrap;
	private HashMap<String, Border> borderMap;
	
	/*
	 * Constructors
	 */
	
	public Style(String name) {
		this.name = name;
		this.valign = "";
		this.halign = "";
		this.format = "";
		this.pattern = "";
		this.wrap = false;
		this.borderMap = new HashMap<String, Border>();
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
	
	public String getFormat() {
		return this.format;
	}
	
	public String getPattern() {
		return this.pattern;
	}
	
	public Boolean getWrap() {
		return this.wrap;
	}
	
	public HashMap<String, Border> getBorderMap() {
		return this.borderMap;
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
	
	public void setFormat(String format) {
		this.format = format;
	}
	
	public void setPattern(String pattern) {
		this.pattern = pattern;
	}
	
	public void setWrap(Boolean wrap) {
		this.wrap = wrap;
	}
	
	public void setBorderMap(HashMap<String, Border> borderMap) {
		this.borderMap = borderMap;
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
