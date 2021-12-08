package com.informatica.xml2xlsx;

public class Validation {
	
	/*
	 * Attribute declaration
	 */
	
	private String name, formula, type, dateExact, dateMin, dateMax;
	private int lengthMin, lengthMax, lengthValue, operator;
	private Float numMin, numMax, numExact;
	private String[] values;
	
	/*
	 * Constructors
	 */
	
	public Validation(String name, String type) {
		this.name = name;
		this.formula = "";
		this.type = type;
		this.values = null;
		this.operator = -1;
		this.lengthValue = -1;
		this.lengthMin = -1;
		this.lengthMax = -1;
		this.numExact = null;
		this.numMin = null;
		this.numMax = null;
		this.dateExact = null;
		this.dateMin = null;
		this.dateMax = null;
	}
	
	/*
	 * Helpers
	 */
	
	private String getDateFuncStr(String date) {
		
		if(date == null) {
			return null;
		}
		else {
			String[] tokens = date.split("-");
			Integer year = Integer.parseInt(tokens[0]);
			Integer month = Integer.parseInt(tokens[1]);
			Integer day = Integer.parseInt(tokens[2]);
			return String.format("Date(%s, %s, %s)", year.toString(), month.toString(), day.toString());
		}
		
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
	
	public int getOperator() {
		return this.operator;
	}
	
	public int getLengthMin() {
		return this.lengthMin;
	}
	
	public int getLengthMax() {
		return this.lengthMax;
	}
	
	public int getLengthValue() {
		return this.lengthValue;
	}
	
	public Float getNumMin() {
		return this.numMin;
	}
	
	public Float getNumMax() {
		return this.numMax;
	}
	
	public Float getNumExact() {
		return this.numExact;
	}
	
	public String getDateMin() {
		return this.dateMin;
	}
	
	public String getDateMax() {
		return this.dateMax;
	}
	
	public String getDateExact() {
		return this.dateExact;
	}
	
	public String getDateMinFunc() {
		return getDateFuncStr(this.dateMin);
	}
	
	public String getDateMaxFunc() {
		return getDateFuncStr(this.dateMax);
	}
	
	public String getDateExactFunc() {
		return getDateFuncStr(this.dateExact);
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
	
	public void setOperator(int operator) {
		this.operator = operator;
	}
	
	public void setLengthMin(int lengthMin) {
		this.lengthMin = lengthMin;
	}
	
	public void setLengthMax(int lengthMax) {
		this.lengthMax = lengthMax;
	}
	
	public void setLengthValue(int lengthValue) {
		this.lengthValue = lengthValue;
	}
	
	public void setNumMin(Float numMin) {
		this.numMin = numMin;
	}
	
	public void setNumMax(Float numMax) {
		this.numMax = numMax;
	}
	
	public void setNumExact(Float numExact) {
		this.numExact = numExact;
	}
	
	public void setDateMin(String dateMin) {
		this.dateMin = dateMin;
	}
	
	public void setDateMax(String dateMax) {
		this.dateMax = dateMax;
	}
	
	public void setDateExact(String dateExact) {
		this.dateExact = dateExact;
	}
	
}
