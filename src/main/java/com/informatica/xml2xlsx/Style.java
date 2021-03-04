package com.informatica.xml2xlsx;

public class Style {
	
	/*
	 * Attribute declaration
	 */
	
	private String name, valign, halign;
	
	/*
	 * Constructors
	 */
	
	public Style(String name) {
		this.name = name;
		this.valign = "";
		this.halign = "";
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
	
}
