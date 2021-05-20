package com.informatica.xml2xlsx;

public class StyleFormat {
	
	// Variable declaration
	private String type = "";
	private String pattern = "";
	
	/*
	 * Constructors
	 */
	
	public StyleFormat(String type) {
		this.type = type;
		this.pattern = "";
	}
	
	public StyleFormat(String type, String pattern) {
		this.type = type;
		this.pattern = pattern;
	}
	
	/*
	 * Getters
	 */
	
	public String getType() {
		return this.type;
	}
	
	public String getPattern() {
		return this.pattern;
	}
	
	/*
	 * Setters
	 */
	
	public void setType(String type) {
		this.type = type;
	}
	
	public void setPattern(String pattern) {
		this.pattern = pattern;
	}
	
	
	
}
