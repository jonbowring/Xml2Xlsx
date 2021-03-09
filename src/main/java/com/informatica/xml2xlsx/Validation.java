package com.informatica.xml2xlsx;

public class Validation {
	
	/*
	 * Attribute declaration
	 */
	
	private String name, formula, type;
	private String[] values;
	
	/*
	 * Constructors
	 */
	
	public Validation(String name, String type) {
		this.name = name;
		this.formula = "";
		this.type = type;
		this.values = null;
	}
	
	/*
	 * Getters
	 */
	
	public String getName() {
		return this.name;
	}
	
	public String getFormula() {
		return this.formula;
	}
	
	public String getType() {
		return this.type;
	}
	
	public String[] getValues() {
		return this.values;
	}
	
	/*
	 * Setters
	 */
	
	public void setName(String name) {
		this.name = name;
	}
	
	public void setFormula(String formula) {
		this.formula = formula;
	}
	
	public void setType(String type) {
		this.type = type;
	}
	
	public void setValues(String[] values) {
		this.values = values;
	}
	
}
