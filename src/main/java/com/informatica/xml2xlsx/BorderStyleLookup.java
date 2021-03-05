package com.informatica.xml2xlsx;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class BorderStyleLookup {
	
	/*
	 * Attribute declaration
	 */
	
	private HashMap<String, BorderStyle> styleMap;
	
	/*
	 * Constructors
	 */
	
	public BorderStyleLookup() {
		
		this.styleMap = new HashMap<String, BorderStyle>();
		this.styleMap.put("dash-dot", BorderStyle.DASH_DOT);
		this.styleMap.put("dash-dot-dot", BorderStyle.DASH_DOT_DOT);
		this.styleMap.put("dashed", BorderStyle.DASHED);
		this.styleMap.put("dotted", BorderStyle.DOTTED);
		this.styleMap.put("double", BorderStyle.DOUBLE);
		this.styleMap.put("hair", BorderStyle.HAIR);
		this.styleMap.put("medium", BorderStyle.MEDIUM);
		this.styleMap.put("medium-dash-dot", BorderStyle.MEDIUM_DASH_DOT);
		this.styleMap.put("medium-dash-dot-dot", BorderStyle.MEDIUM_DASH_DOT_DOT);
		this.styleMap.put("medium-dashed", BorderStyle.MEDIUM_DASHED);
		this.styleMap.put("none", BorderStyle.NONE);
		this.styleMap.put("slanted-dash-dot", BorderStyle.SLANTED_DASH_DOT);
		this.styleMap.put("thick", BorderStyle.THICK);
		this.styleMap.put("thin", BorderStyle.THIN);
		
	}
	
	/*
	 * Getters
	 */
	
	public HashMap<String, BorderStyle> getBorderStyles() {
		return this.styleMap;
	}
	
	
}
