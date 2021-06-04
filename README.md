# Xml2Xlsx

## Description:
Xml2Xlsx is an executable jar that enables you to create Excel XLSX files using a simple XML markup. It provides the ability for managing multiple worksheets, formatting, styling, data validation and tables.

## Usage:
```
java -jar Xml2Xlsx-[Version].jar --src [Source XML file] --tgt [Target Excel file] --showProgress
```

### Example Usage:
```
java -jar xml2xlsx-1.2.2.jar --src "books.xml" --tgt "books.xlsx" --showProgress
```

## Command Line Options:
Option|Description
------|-----------
--src|Used to specify the location of the input XML file.
--tgt|Used to specify the location of the output Excel file.
--showProgress|Optional. Used to display a progress bar when writing rows to the target.

## XML Markup:
### Styles Markup:
XPath|Description
-----|-----------
/workbook/styles/style|Optional. Used to define re-usable styles to be applied to the cells.
/workbook/styles/style/@name|Mandatory. The name for the style. This name is used by cells to reference the style to be applied. The name can only contain numbers, letters and underscores.
/workbook/styles/style/align|Optional. Used to define the horizontal and vertical alignment for a style.
/workbook/styles/style/align/@vertical|Optional. Used to define the vertical alignment for an align property. Possible values include "top", "center" and "bottom".
/workbook/styles/style/align/@horizontal|Optional. Used to define the horizontal alignment for an align property. Possible values include "left", "center" and "right".
/workbook/styles/style/border|Optional, can have up to 4 borders defined. Used to define the border style for a cell.
/workbook/styles/style/border/@pos|Mandatory. Used to define which side of the cell the border will be applied to. Possible values include "top", "right", "bottom" and "left".
/workbook/styles/style/border/@type|Optional. Used to define the line style of the border. Possible values include "dash-dot", "dash-dot-dot", "dashed", "dotted", "double", "hair", "medium", "medium-dash-dot", "medium-dash-dot-dot", "medium-dashed", "none", "slanted-dash-dot", "thick" and "thin". If not defined the type "thin" will be applied.
/workbook/styles/style/border/@colour|Optional. Used to define the colour of the border. The colour can be defined as either an rgb colour using the format "rgb([red],[green],[blue])" (for example "rgb(125,36,210)") or using a pre-defined colour label. Possible pre-defined colour labels include "aqua", "automatic", "black", "black1", "blue", "blue1", "blue-grey", "bright-green", "bright-green1", "brown", "coral", "cornflower-blue", "dark-blue", "dark-green", "dark-red", "dark-teal", "dark-yellow", "gold", "green", "grey-25-percent", "grey-40-percent", "grey-50-percent", "grey-80-percent", "indigo", "lavender", "lemon-chiffon", "light-blue", "light-cornflower-blue", "light-green", "light-orange", "light-turquoise", "light-turquoise1", "light-yellow", "lime", "maroon", "olive-green", "orange", "orchid", "pale-blue", "pink", "pink1", "plum", "red", "red1", "rose", "royal-blue", "sea-green", "sky-blue", "tan", "tan", "turquoise", "turquoise1", "violet", "white", "white1", "yellow" and "yellow1".
/workbook/styles/style/fill|Optional. Used to define the fill style for a cell.
/workbook/styles/style/fill/@colour|Mandatory. Used to define the colour of the fill. The colour can be defined as either an rgb colour using the format "rgb([red],[green],[blue])" (for example "rgb(125,36,210)") or using a pre-defined colour label. Possible pre-defined colour labels include "aqua", "automatic", "black", "black1", "blue", "blue1", "blue-grey", "bright-green", "bright-green1", "brown", "coral", "cornflower-blue", "dark-blue", "dark-green", "dark-red", "dark-teal", "dark-yellow", "gold", "green", "grey-25-percent", "grey-40-percent", "grey-50-percent", "grey-80-percent", "indigo", "lavender", "lemon-chiffon", "light-blue", "light-cornflower-blue", "light-green", "light-orange", "light-turquoise", "light-turquoise1", "light-yellow", "lime", "maroon", "olive-green", "orange", "orchid", "pale-blue", "pink", "pink1", "plum", "red", "red1", "rose", "royal-blue", "sea-green", "sky-blue", "tan", "tan", "turquoise", "turquoise1", "violet", "white", "white1", "yellow" and "yellow1".
/workbook/styles/style/fill/@pattern|Optional. Used to define the fill pattern of the cell. Possible values include "alt-bars", "big-spots", "bricks", "diamonds", "fine-dots", "least-dots", "less-dots", "no-fill", "solid-foreground", "sparse-dots", "squares", "thick-backward-diag", "thick-forward-diag", "thick-horz-bands", "thick-vert-bands", "thin-backward-diag", "thin-forward-diag", "thin-horz-bands" and "thin-vert-bands".
/workbook/styles/style/font|Optional. Used to define the font styling of the cell.
/workbook/styles/style/font/@name|Mandatory. The name of the font style to be applied. The value should match the font names used by the operating system.
/workbook/styles/style/font/@size|Optional. An integer used to set the font size in points.
/workbook/styles/style/font/@colour|Optional. Used to define the colour of the font. The colour can be defined as either an rgb colour using the format "rgb([red],[green],[blue])" (for example "rgb(125,36,210)") or using a pre-defined colour label. Possible pre-defined colour labels include "aqua", "automatic", "black", "black1", "blue", "blue1", "blue-grey", "bright-green", "bright-green1", "brown", "coral", "cornflower-blue", "dark-blue", "dark-green", "dark-red", "dark-teal", "dark-yellow", "gold", "green", "grey-25-percent", "grey-40-percent", "grey-50-percent", "grey-80-percent", "indigo", "lavender", "lemon-chiffon", "light-blue", "light-cornflower-blue", "light-green", "light-orange", "light-turquoise", "light-turquoise1", "light-yellow", "lime", "maroon", "olive-green", "orange", "orchid", "pale-blue", "pink", "pink1", "plum", "red", "red1", "rose", "royal-blue", "sea-green", "sky-blue", "tan", "tan", "turquoise", "turquoise1", "violet", "white", "white1", "yellow" and "yellow1".
/workbook/styles/style/font/italic|Optional. An empty element used as a flag to indicate if the font should have italic styling applied.
/workbook/styles/style/font/strikeout|Optional. An empty element used as a flag to indicate if the font should have the strikeout styling applied.
/workbook/styles/style/wrap|Optional. An empty element used as a flag to indicate text wrapping should be applied to the cell. If a cell contains newline characters represented as "\n" then this flag must be included for the newlines to be properly displayed.
/workbook/styles/style/format|Optional. Used to define the data type and pattern format applied to the cell.
/workbook/styles/style/format/@type|Mandatory. Used to specify the data type. Possible values include "currency", "date", "datetime", "float", "formula", "fraction", "int", "percent", "scientific" and "string". Note: when a cell uses the format "date" the XML value must be in the format "yyyy-MM-dd". When a cell uses the format "datetime" the XML value must be in the format "yyyy-MM-dd hh:mm:ss".
/workbook/styles/style/format/@pattern|Optional. If @type is specified as a "currency", "date", "datetime" or "percent" then this attribute can be used to define a custom Excel pattern (e.g. "dd/MM/yyyy"). If the pattern is not included then it will default to the Excel default format.
/workbook/styles/style/format/@separator|Optional. If @type is specified as a "float" or "int" then this attribute can be set to "true" to include the thousands separator.

### Data Validations Markup:
XPath|Description
-----|-----------
/workbook/validations/validation|Optional. Used to define re-usable data validation rules to be applied to cells.
/workbook/validations/validation/@name|Mandatory. The name for the data validation. This name is used by cells to reference the data validation to be applied. The name can only contain numbers, letters and underscores.
/workbook/validations/validation/type|Mandatory. The type of data validation to be applied. Possible values include "fixed-list" and "formula-list".
/workbook/validations/validation/values|Mandatory if type is set to "fixed-list".
/workbook/validations/validation/values/value|Mandatory. One or more values to be used in the data validation.
/workbook/validations/validation/formula|Mandatory if type is set to "formula-list". Can be specified using Excel style reference formulas, including other tabs. For example "'Books'!$B$2:$B$5".

### Worksheet Markup:
XPath|Description
-----|-----------
/workbook/worksheet|Mandatory. Used to specify a worksheet tab to be included in the Excel file.
/workbook/worksheet/@name|Mandatory. The name of the worksheet tab.
/workbook/worksheet/@autofilter|Optional. Used to define if auto filters should be applied to the first row in the worksheet. This option is ignored if a table has been defined for the worksheet as auto filters are automatically applied to tables. Possible values are "true" or "false".
/workbook/worksheet/@autofit|Optional. Used to define if the columns for the worksheet should automatically be resized to fit the contents. Possible values are "true" or "false".
/workbook/worksheet/@hidden|Optional. Used to set the visibility of the worksheet. IMPORTANT: There must always e at least one worksheet visible. Possible values are "true" or "false".
/workbook/worksheet/table|Optional. Used to define if the worksheet data should be contained within a table.
/workbook/worksheet/table/@name|Mandatory. The name for the table. The name can only contain numbers, letters and underscores.
/workbook/worksheet/table/@colStripes|Optional. Used to specify if column colour striping should be applied. Possible values include "true" and "false".
/workbook/worksheet/table/@rowStripes|Optional. Used to specify if row colour striping should be applied. Possible values include "true" and "false".
/workbook/worksheet/table/@style|Optional. Used to define the style type of the table. Possible values include "TableStyleDark1", "TableStyleDark2", "TableStyleDark3", "TableStyleDark4", "TableStyleDark5", "TableStyleDark6", "TableStyleDark7", "TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleDark21", "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21", "TableStyleMedium22", "TableStyleMedium23", "TableStyleMedium24", "TableStyleMedium25", "TableStyleMedium26", "TableStyleMedium27" and "TableStyleMedium28".
/workbook/worksheet/row|Mandatory. Used to specify a row of data to be added to the Excel file. Maximum number of rows that can be included is 1,048,576.
/workbook/worksheet/row/cell|Mandatory. Used to specify a cell of data to be added to the Excel file. Maximum number of cells or columns that can be included is 16,384.
/workbook/worksheet/row/cell/@style|Optional. The name of the re-usable style to be applied to the cell.
/workbook/worksheet/row/cell/@validation|Optional. The name of the re-usable validation to be applied to the cell.


## Examples:
### A Simple Worksheet With Auto Filters:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example1.png?raw=true" alt="A Simple Worksheet With Auto Filters"/>

```
<workbook>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell>Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell>2005</cell>
			<cell>30.00</cell>
		</row>
		<row>
			<cell>Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
		</row>
		<row>
			<cell>XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
		</row>
		<row>
			<cell>Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Cell Data Types:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example2.png?raw=true" alt="Cell Data Types"/>

```
<workbook>
	<styles>
		<style name="myInt">
			<format type="int"/>
		</style>
		<style name="myFloat">
			<format type="float"/>
		</style>
	</styles>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell>Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell style="myInt">2005</cell>
			<cell style="myFloat">30.00</cell>
		</row>
		<row>
			<cell>Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell style="myInt">2005</cell>
			<cell style="myFloat">29.99</cell>
		</row>
		<row>
			<cell>XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell style="myInt">2003</cell>
			<cell style="myFloat">49.99</cell>
		</row>
		<row>
			<cell>Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell style="myInt">2003</cell>
			<cell style="myFloat">39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Formulas:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example3.png?raw=true" alt="Formulas"/>

```
<workbook>
	<styles>
		<style name="myFormula">
			<format type="formula"/>
		</style>
	</styles>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
			<cell>Formula</cell>
		</row>
		<row>
			<cell>Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell>2005</cell>
			<cell>30.00</cell>
			<cell style="myFormula">A2&amp;" - "&amp;B2</cell>
		</row>
		<row>
			<cell>Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
			<cell style="myFormula">C2 * D2</cell>
		</row>
		<row>
			<cell>XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
			<cell style="myFormula">upper(A2)</cell>
		</row>
		<row>
			<cell>Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
			<cell style="myFormula">lower(B2)</cell>
		</row>
	</worksheet>
</workbook>
```

### Basic Table:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example4.png?raw=true" alt="Basic Table"/>

```
<workbook>
	<worksheet name="Books" autofilter="true">
		<table name="My_Table" colStripes="false" rowStripes="true" style="TableStyleMedium3"/>
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell>Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell>2005</cell>
			<cell>30.00</cell>
		</row>
		<row>
			<cell>Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
		</row>
		<row>
			<cell>XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
		</row>
		<row>
			<cell>Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Cell Borders:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example5.png?raw=true" alt="Cell Borders"/>

```
<workbook>
	<styles>
		<style name="myBorders">
			<border pos="top" type="medium-dashed" colour="red"/>
			<border pos="right" type="medium-dashed" colour="red"/>
			<border pos="bottom" type="medium-dashed" colour="red"/>
			<border pos="left" type="medium-dashed" colour="red"/>
		</style>
	</styles>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell>Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell style="myBorders">2005</cell>
			<cell>30.00</cell>
		</row>
		<row>
			<cell>Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell style="myBorders">2005</cell>
			<cell>29.99</cell>
		</row>
		<row>
			<cell>XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell style="myBorders">2003</cell>
			<cell>49.99</cell>
		</row>
		<row>
			<cell>Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell style="myBorders">2003</cell>
			<cell>39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Fonts:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example6.png?raw=true" alt="Fonts"/>

```
<workbook>
	<styles>
		<style name="myFont">
			<font name="Courier New" size="24" colour="blue">
				<italic/>
				<strikeout/>
			</font>
		</style>
	</styles>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell style="myFont">Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell>2005</cell>
			<cell>30.00</cell>
		</row>
		<row>
			<cell style="myFont">Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
		</row>
		<row>
			<cell style="myFont">XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
		</row>
		<row>
			<cell style="myFont">Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Fill:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example7.png?raw=true" alt="Fonts"/>

```
<workbook>
	<styles>
		<style name="myFill">
			<fill colour="rgb(100,200,50)" pattern="squares"/>
		</style>
	</styles>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell style="myFill">Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell>2005</cell>
			<cell>30.00</cell>
		</row>
		<row>
			<cell style="myFill">Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
		</row>
		<row>
			<cell style="myFill">XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
		</row>
		<row>
			<cell style="myFill">Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Data Validations:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example8.png?raw=true" alt="Fonts"/>

```
<workbook>
	<validations>
		<validation name="my_validation1">
			<type>fixed-list</type>
			<values>
				<value>FOO</value>
				<value>BAR</value>
				<value>CAT</value>
			</values>
		</validation>
		<validation name="my_validation2">
			<type>formula-list</type>
			<formula>'Books'!$B$2:$B$5</formula>
		</validation>
	</validations>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell validation="my_validation1">Everyday Italian</cell>
			<cell validation="my_validation2">Giada De Laurentiis</cell>
			<cell>2005</cell>
			<cell>30.00</cell>
		</row>
		<row>
			<cell validation="my_validation1">Harry Potter</cell>
			<cell validation="my_validation2">J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
		</row>
		<row>
			<cell validation="my_validation1">XQuery Kick Start</cell>
			<cell validation="my_validation2">Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
		</row>
		<row>
			<cell validation="my_validation1">Learning XML</cell>
			<cell validation="my_validation2">Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Mash up:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example9.png?raw=true" alt="Fonts"/>

```
<workbook>
	<styles>
		<style name="top_left">
			<align horizontal="left" vertical="top"/>
			<font name="Courier New" size="24" colour="blue">
				<italic/>
				<strikeout/>
			</font>
		</style>
		<style name="bottom_right">
			<align horizontal="right" vertical="bottom"/>
			<wrap/>
		</style>
		<style name="center">
			<align horizontal="center" vertical="center"/>
			<border pos="top" type="slanted-dash-dot" colour="rgb(100,200,50)"/>
			<border pos="right" type="dashed" colour="aqua"/>
			<border pos="bottom" type="thick" colour="rgb(100,200,50)"/>
			<border pos="left" type="double" colour="rgb(100,200,50)"/>
			<font name="Courier New">
				<size>50</size>
			</font>
		</style>
		<style name="formula">
			<format type="formula"/>
		</style>
		<style name="date">
			<format type="date" pattern="dd/MM/yyyy"/>
		</style>
		<style name="float">
			<format type="float"/>
			<fill colour="aqua"/>
		</style>
		<style name="int">
			<fill colour="rgb(100,200,50)" pattern="squares"/>
			<format type="int"/>
		</style>
	</styles>
	<validations>
		<validation name="my_validation1">
			<type>fixed-list</type>
			<values>
				<value>FOO</value>
				<value>BAR</value>
				<value>CAT</value>
			</values>
		</validation>
		<validation name="my_validation2">
			<type>formula-list</type>
			<formula>'Books'!$B$2:$B$5</formula>
		</validation>
	</validations>
	<worksheet name="Books" autofilter="false">
		<table name="My_Table" colStripes="false" rowStripes="true" style="TableStyleMedium2"/>
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell style="formula"  validation="my_validation1">1+1</cell>
			<cell>Giada De Laurentiis</cell>
			<cell style="date">2005-11-26 00:00:01</cell>
			<cell style="float">30.00</cell>
		</row>
		<row>
			<cell style="bottom_right">Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</cell>
			<cell style="top_left">J K. Rowling</cell>
			<cell style="int">2005</cell>
			<cell style="float">29.99</cell>
		</row>
		<row>
			<cell validation="my_validation1">XQuery Kick Start</cell>
			<cell>James McGovern</cell>
			<cell style="center">2003</cell>
			<cell style="float">49.99</cell>
		</row>
		<row>
			<cell validation="my_validation2">Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell style="int">2003</cell>
			<cell style="float">39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### A Simple Workbook With A Hidden Worksheet:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example10.png?raw=true" alt="A Simple Workbook With A Hidden Worksheet"/>

```
<workbook>
	<worksheet name="Books">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
		</row>
		<row>
			<cell>Everyday Italian</cell>
			<cell>Giada De Laurentiis</cell>
			<cell>2005</cell>
			<cell>30.00</cell>
		</row>
		<row>
			<cell>Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
		</row>
		<row>
			<cell>XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
		</row>
		<row>
			<cell>Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
		</row>
	</worksheet>
	<worksheet name="Secret" hidden="true">
		<row>
			<cell>This worksheet is not visible.</cell>
		</row>
	</worksheet>
</workbook>
```