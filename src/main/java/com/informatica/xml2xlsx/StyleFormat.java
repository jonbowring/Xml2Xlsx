package com.informatica.xml2xlsx;

public class StyleFormat {
	
	// Attribute declaration
	String type;
	Boolean isFormula;
	
	/*
	 * -----------------------
	 * Constructors
	 * -----------------------
	 */
	
	public StyleFormat(String type) {
		this.type = type;
		this.isFormula = false;
	}
	
	public StyleFormat(String type, Boolean isFormula) {
		this.type = type;
		this.isFormula = isFormula;
	}
	
	/*
	 * -----------------------
	 * Getters
	 * -----------------------
	 */
	
	public String getType() {
		return this.type;
	}
	
	public Boolean getIsFormula() {
		return this.isFormula;
	}
	
	/*
	 * -----------------------
	 * Setters
	 * -----------------------
	 */
	
	public void setType(String type) {
		this.type = type;
	}
	
	public void setIsFormula(Boolean isFormula) {
		this.isFormula = isFormula;
	}
	
}
