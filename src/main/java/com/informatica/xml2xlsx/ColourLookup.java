package com.informatica.xml2xlsx;

import java.util.HashMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

public class ColourLookup {
	
	/*
	 * Attribute declaration
	 */
	
	private HashMap<String, IndexedColors> colourMap;
	
	/*
	 * Constructors
	 */
	
	public ColourLookup() {
		
		this.colourMap = new HashMap<String, IndexedColors>();
		this.colourMap.put("aqua", IndexedColors.AQUA);
		this.colourMap.put("automatic", IndexedColors.AUTOMATIC);
		this.colourMap.put("black", IndexedColors.BLACK);
		this.colourMap.put("black1", IndexedColors.BLACK1);
		this.colourMap.put("blue", IndexedColors.BLUE);
		this.colourMap.put("blue1", IndexedColors.BLUE1);
		this.colourMap.put("blue-grey", IndexedColors.BLUE_GREY);
		this.colourMap.put("bright-green", IndexedColors.BRIGHT_GREEN);
		this.colourMap.put("bright-green1", IndexedColors.BRIGHT_GREEN1);
		this.colourMap.put("brown", IndexedColors.BROWN);
		this.colourMap.put("coral", IndexedColors.CORAL);
		this.colourMap.put("cornflower-blue", IndexedColors.CORNFLOWER_BLUE);
		this.colourMap.put("dark-blue", IndexedColors.DARK_BLUE);
		this.colourMap.put("dark-green", IndexedColors.DARK_GREEN);
		this.colourMap.put("dark-red", IndexedColors.DARK_RED);
		this.colourMap.put("dark-teal", IndexedColors.DARK_TEAL);
		this.colourMap.put("dark-yellow", IndexedColors.DARK_YELLOW);
		this.colourMap.put("gold", IndexedColors.GOLD);
		this.colourMap.put("green", IndexedColors.GREEN);
		this.colourMap.put("grey-25-percent", IndexedColors.GREY_25_PERCENT);
		this.colourMap.put("grey-40-percent", IndexedColors.GREY_40_PERCENT);
		this.colourMap.put("grey-50-percent", IndexedColors.GREY_50_PERCENT);
		this.colourMap.put("grey-80-percent", IndexedColors.GREY_80_PERCENT);
		this.colourMap.put("indigo", IndexedColors.INDIGO);
		this.colourMap.put("lavender", IndexedColors.LAVENDER);
		this.colourMap.put("lemon-chiffon", IndexedColors.LEMON_CHIFFON);
		this.colourMap.put("light-blue", IndexedColors.LIGHT_BLUE);
		this.colourMap.put("light-cornflower-blue", IndexedColors.LIGHT_CORNFLOWER_BLUE);
		this.colourMap.put("light-green", IndexedColors.LIGHT_GREEN);
		this.colourMap.put("light-orange", IndexedColors.LIGHT_ORANGE);
		this.colourMap.put("light-turquoise", IndexedColors.LIGHT_TURQUOISE);
		this.colourMap.put("light-turquoise1", IndexedColors.LIGHT_TURQUOISE1);
		this.colourMap.put("light-yellow", IndexedColors.LIGHT_YELLOW);
		this.colourMap.put("lime", IndexedColors.LIME);
		this.colourMap.put("maroon", IndexedColors.MAROON);
		this.colourMap.put("olive-green", IndexedColors.OLIVE_GREEN);
		this.colourMap.put("orange", IndexedColors.ORANGE);
		this.colourMap.put("orchid", IndexedColors.ORCHID);
		this.colourMap.put("pale-blue", IndexedColors.PALE_BLUE);
		this.colourMap.put("pink", IndexedColors.PINK);
		this.colourMap.put("pink1", IndexedColors.PINK1);
		this.colourMap.put("plum", IndexedColors.PLUM);
		this.colourMap.put("red", IndexedColors.RED);
		this.colourMap.put("red1", IndexedColors.RED1);
		this.colourMap.put("rose", IndexedColors.ROSE);
		this.colourMap.put("royal-blue", IndexedColors.ROYAL_BLUE);
		this.colourMap.put("sea-green", IndexedColors.SEA_GREEN);
		this.colourMap.put("sky-blue", IndexedColors.SKY_BLUE);
		this.colourMap.put("tan", IndexedColors.TAN);
		this.colourMap.put("tan", IndexedColors.TAN);
		this.colourMap.put("turquoise", IndexedColors.TURQUOISE);
		this.colourMap.put("turquoise1", IndexedColors.TURQUOISE1);
		this.colourMap.put("violet", IndexedColors.VIOLET);
		this.colourMap.put("white", IndexedColors.WHITE);
		this.colourMap.put("white1", IndexedColors.WHITE1);
		this.colourMap.put("yellow", IndexedColors.YELLOW);
		this.colourMap.put("yellow1", IndexedColors.YELLOW1);
		
	}
	
	/*
	 * Getters
	 */
	
	public HashMap<String, IndexedColors> getColours() {
		return this.colourMap;
	}
	
}
