package com.informatica.xml2xlsx;

public class Border {
	
	/*
	 * Attribute declaration
	 */
	
	private String pos, type, colour;
	
	/*
	 * Constructors
	 */
	
	public Border(String pos) {
		this.pos = pos;
		this.type = "thin";
		this.colour = "black";
	}
	
	public Border(String pos, String type) {
		this.pos = pos;
		this.type = type;
		this.colour = "black";
	}
	
	public Border(String pos, String type, String colour) {
		this.pos = pos;
		this.type = type;
		this.colour = colour;
	}
	
	/*
	 * Getters
	 */
	
	public String getPos() {
		return this.pos;
	}
	
	public String getType() {
		return this.type;
	}
	
	public String getColour() {
		return this.colour;
	}
	
	/*
	 * Setters
	 */
	
	public void setPos(String pos) {
		this.pos = pos;
	}
	
	public void setType(String type) {
		this.type = type;
	}
	
	public void setColour(String colour) {
		this.colour = colour;
	}
	
}
