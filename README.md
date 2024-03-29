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
/workbook/styles/style/font/@name|Optional. The name of the font style to be applied. The value should match the font names used by the operating system. If not set then the default font is used.
/workbook/styles/style/font/@size|Optional. An integer used to set the font size in points.
/workbook/styles/style/font/@colour|Optional. Used to define the colour of the font. The colour can be defined as either an rgb colour using the format "rgb([red],[green],[blue])" (for example "rgb(125,36,210)") or using a pre-defined colour label. Possible pre-defined colour labels include "aqua", "automatic", "black", "black1", "blue", "blue1", "blue-grey", "bright-green", "bright-green1", "brown", "coral", "cornflower-blue", "dark-blue", "dark-green", "dark-red", "dark-teal", "dark-yellow", "gold", "green", "grey-25-percent", "grey-40-percent", "grey-50-percent", "grey-80-percent", "indigo", "lavender", "lemon-chiffon", "light-blue", "light-cornflower-blue", "light-green", "light-orange", "light-turquoise", "light-turquoise1", "light-yellow", "lime", "maroon", "olive-green", "orange", "orchid", "pale-blue", "pink", "pink1", "plum", "red", "red1", "rose", "royal-blue", "sea-green", "sky-blue", "tan", "tan", "turquoise", "turquoise1", "violet", "white", "white1", "yellow" and "yellow1".
/workbook/styles/style/font/italic|Optional. An empty element used as a flag to indicate if the font should have italic styling applied.
/workbook/styles/style/font/strikeout|Optional. An empty element used as a flag to indicate if the font should have the strikeout styling applied.
/workbook/styles/style/font/bold|Optional. An empty element used as a flag to indicate if the font should be bold.
/workbook/styles/style/font/underline|Optional. An empty element used as a flag to indicate if the font should be underlined.
/workbook/styles/style/wrap|Optional. An empty element used as a flag to indicate text wrapping should be applied to the cell. If a cell contains newline characters represented as "&#10;" then this flag must be included for the newlines to be properly displayed.
/workbook/styles/style/format|Optional. Used to define the data type and pattern format applied to the cell.
/workbook/styles/style/format/@type|Mandatory. Used to specify the data type. Possible values include "currency", "date", "datetime", "float", "fraction", "general", "int", "percent", "scientific" and "string". Note: when a cell uses the format "date" the XML value must be in the format "yyyy-MM-dd". When a cell uses the format "datetime" the XML value must be in the format "yyyy-MM-dd hh:mm:ss".
/workbook/styles/style/format/@formula|Optional. Used to indicate if the cell value should be treated as a formula. Possible values include "true" or "false".
/workbook/styles/style/format/@pattern|Optional. If @type is specified as a "currency", "date", "datetime" or "percent" then this attribute can be used to define a custom Excel pattern (e.g. "dd/MM/yyyy"). If the pattern is not included then it will default to the Excel default format.
/workbook/styles/style/format/@separator|Optional. If @type is specified as a "float" or "int" then this attribute can be set to "true" to include the thousands separator.

### Data Validations Markup:
XPath|Description
-----|-----------
/workbook/validations/validation|Optional. Used to define re-usable data validation rules to be applied to cells.
/workbook/validations/validation/@name|Mandatory. The name for the data validation. This name is used by cells to reference the data validation to be applied. The name can only contain numbers, letters and underscores.
/workbook/validations/validation/type|Mandatory. The type of data validation to be applied. Possible values include "fixed-list", "formula-list", "length", "numerical" and "date".
/workbook/validations/validation/value|Mandatory if type is set to "length", "numerical" or "date" and the "operator" is either "EQUAL", "NOT_EQUAL", "GREATER_THAN", "GREATER_OR_EQUAL", "LESS_THAN" or "LESS_OR_EQUAL". Date values must be entered using the format "yyyy-MM-dd". Used to determine the operator to be applied for the "length" and "value" validations.
/workbook/validations/validation/min|Mandatory if type is set to "length", "numerical" or "date" and the "operator" is either "BETWEEN" or "NOT_BETWEEN". Date values must be entered using the format "yyyy-MM-dd". Used to determine the minimum length value for the validation to be applied.
/workbook/validations/validation/max|Mandatory if type is set to "length", "numerical" or "date" and the "operator" is either "BETWEEN" or "NOT_BETWEEN". Date values must be entered using the format "yyyy-MM-dd". Used to determine the maximum length value for the validation to be applied.
/workbook/validations/validation/operator|Mandatory if type is set to "length" or "numerical". Possible operations include "EQUAL", "NOT_EQUAL", "GREATER_THAN", "GREATER_OR_EQUAL", "LESS_THAN", "LESS_OR_EQUAL", "BETWEEN" and "NOT_BETWEEN".
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
/workbook/worksheet/@gridlines|Optional. Used to define if gridlines are displayed for the worksheet. Possible values are "true" or "false".
/workbook/worksheet/@hidden|Optional. Used to set the visibility of the worksheet. IMPORTANT: There must always e at least one worksheet visible. Possible values are "true" or "false".
/workbook/worksheet/@order|Optional. The index of the position for the worksheet tab to be displayed. The first tab has an index of 0.
/workbook/worksheet/columns|Optional. Used to define settings for columns within a worksheet.
/workbook/worksheet/column|Mandatory. Used to define column level settings if needed.
/workbook/worksheet/column/@index|Mandatory. The index of the column for the settings to be applied to. The first column has an index of 0.
/workbook/worksheet/column/@width|Optional. Sets the width (in units of 1/256th of a character width).
/workbook/worksheet/column/@style|Optional. Sets the default style for a column.
/workbook/worksheet/table|Optional. Used to define if the worksheet data should be contained within a table.
/workbook/worksheet/table/@name|Mandatory. The name for the table. The name can only contain numbers, letters and underscores.
/workbook/worksheet/table/@colStripes|Optional. Used to specify if column colour striping should be applied. Possible values include "true" and "false".
/workbook/worksheet/table/@rowStripes|Optional. Used to specify if row colour striping should be applied. Possible values include "true" and "false".
/workbook/worksheet/table/@style|Optional. Used to define the style type of the table. Possible values include "TableStyleDark1", "TableStyleDark2", "TableStyleDark3", "TableStyleDark4", "TableStyleDark5", "TableStyleDark6", "TableStyleDark7", "TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleDark21", "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21", "TableStyleMedium22", "TableStyleMedium23", "TableStyleMedium24", "TableStyleMedium25", "TableStyleMedium26", "TableStyleMedium27" and "TableStyleMedium28".
/workbook/worksheet/row|Mandatory. Used to specify a row of data to be added to the Excel file. Maximum number of rows that can be included is 1,048,576.
/workbook/worksheet/row/cell|Mandatory. Used to specify a cell of data to be added to the Excel file. Maximum number of cells or columns that can be included is 16,384.
/workbook/worksheet/row/cell/@style|Optional. The name of the re-usable style to be applied to the cell.
/workbook/worksheet/row/cell/@columnStyle|Optional. The name of the re-usable style to be applied to the entire column. This is useful if you'd like to specify a style at a header row level.
/workbook/worksheet/row/cell/@validation|Optional. The name of the re-usable validation to be applied to the cell.
/workbook/worksheet/pivot|Optional. Used to define if a worksheet should include a pivot table.
/workbook/worksheet/pivot/@location|Mandatory. Defines the top left cell for positioning the pivot table (e.g. "A1").
/workbook/worksheet/pivot/@dataSheet|Mandatory. Defines the worksheet name that contains the data the pivot table will refer to.
/workbook/worksheet/pivot/@dataArea|Mandatory if NOT using table data. Defines the area reference for the data that the pivot table will refer to (e.g. "A1:C5").
/workbook/worksheet/pivot/@dataTable|Mandatory if using table data. Defines the table reference using the table name for the data that the pivot table will refer to (e.g. "My_Table").
/workbook/worksheet/pivot/groupby|Mandatory. Contains the columns that will be used for grouping in the pivot table.
/workbook/worksheet/pivot/groupby/column|Mandatory - one or more. The column that will be used for grouping in the pivot table.
/workbook/worksheet/pivot/groupby/column/@index|Mandatory. The zero based index of the column to be used for grouping in the pivot table.
/workbook/worksheet/pivot/aggregate|Mandatory. Contains the columns that will be used for aggregating in the pivot table.
/workbook/worksheet/pivot/aggregate/column|Mandatory - one or more. The column that will be used for aggregating in the pivot table.
/workbook/worksheet/pivot/aggregate/column/@index|Mandatory. The zero based index of the column to be used for aggregating in the pivot table.
/workbook/worksheet/pivot/aggregate/column/@action|Mandatory. The aggregate function to be performed on the column. Possible functions include: "AVERAGE", "COUNT", "COUNT_NUMS", "MAX", "MIN", "PRODUCT", "STD_DEV", "STD_DEVP", "SUM", "VAR" and "VARP".
/workbook/worksheet/pivot/aggregate/column/@name|Optional. If needed you can specify a custom name for the column using this attribute.
/workbook/worksheet/pivot/filter|Optional. Contains the columns that will be used for filtering in the pivot table.
/workbook/worksheet/pivot/filter/column|Mandatory - one or more. The column that will be used for filtering in the pivot table.
/workbook/worksheet/pivot/filter/column/@index|Mandatory. The zero based index of the column to be used for filtering in the pivot table.



## Examples:
### A Simple Worksheet With Auto Filters:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example1.png?raw=true" alt="A Simple Worksheet With Auto Filters"/>

```XML
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

```XML
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

```XML
<workbook>
	<styles>
		<style name="myStringFormula">
			<format type="string" formula="true"/>
		</style>
		<style name="myFloatFormula">
			<format type="float" formula="true"/>
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
			<cell style="myStringFormula">A2&amp;" - "&amp;B2</cell>
		</row>
		<row>
			<cell>Harry Potter</cell>
			<cell>J K. Rowling</cell>
			<cell>2005</cell>
			<cell>29.99</cell>
			<cell style="myFloatFormula">C2 * D2</cell>
		</row>
		<row>
			<cell>XQuery Kick Start</cell>
			<cell>Vaidyanathan Nagarajan</cell>
			<cell>2003</cell>
			<cell>49.99</cell>
			<cell style="myStringFormula">upper(A2)</cell>
		</row>
		<row>
			<cell>Learning XML</cell>
			<cell>Erik T. Ray</cell>
			<cell>2003</cell>
			<cell>39.95</cell>
			<cell style="myStringFormula">lower(B2)</cell>
		</row>
	</worksheet>
</workbook>
```

### Basic Table:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example4.png?raw=true" alt="Basic Table"/>

```XML
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

```XML
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

```XML
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

```XML
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

```XML
<workbook>
  <styles>
		<style name="myInt">
			<format type="int"/>
		</style>
		<style name="myFloat">
			<format type="float"/>
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
		<validation name="my_validation3">
			<type>length</type>
			<operator>GREATER_OR_EQUAL</operator>
			<value>10</value>
		</validation>
		<validation name="my_validation4">
			<type>length</type>
			<operator>LESS_OR_EQUAL</operator>
			<value>10</value>
		</validation>
		<validation name="my_validation5">
			<type>length</type>
			<operator>EQUAL</operator>
			<value>10</value>
		</validation>
		<validation name="my_validation6">
			<type>length</type>
			<operator>BETWEEN</operator>
			<min>10</min>
			<max>20</max>
		</validation>
		<validation name="my_validation7">
			<type>length</type>
			<operator>NOT_BETWEEN</operator>
			<min>10</min>
			<max>20</max>
		</validation>
    <validation name="my_validation8">
			<type>numerical</type>
			<operator>GREATER_OR_EQUAL</operator>
			<value>10</value>
		</validation>
		<validation name="my_validation9">
			<type>numerical</type>
			<operator>LESS_OR_EQUAL</operator>
			<value>10</value>
		</validation>
		<validation name="my_validation10">
			<type>numerical</type>
			<operator>EQUAL</operator>
			<value>10</value>
		</validation>
		<validation name="my_validation11">
			<type>numerical</type>
			<operator>BETWEEN</operator>
			<min>10</min>
			<max>20</max>
		</validation>
		<validation name="my_validation12">
			<type>numerical</type>
			<operator>NOT_BETWEEN</operator>
			<min>10</min>
			<max>20</max>
		</validation>
	</validations>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Title</cell>
			<cell>Author</cell>
			<cell>Year</cell>
			<cell>Price</cell>
			<cell>Comments</cell>
			<cell>Remarks</cell>
			<cell>Addendum</cell>
		</row>
		<row>
			<cell validation="my_validation1">Everyday Italian</cell>
			<cell validation="my_validation2">Giada De Laurentiis</cell>
			<cell style="myInt" validation="my_validation8">2005</cell>
			<cell style="myFloat" validation="my_validation9">30.00</cell>
			<cell validation="my_validation5">Some text</cell>
			<cell validation="my_validation6">More text</cell>
			<cell validation="my_validation7">Even more text!</cell>
		</row>
		<row>
			<cell validation="my_validation1">Harry Potter</cell>
			<cell validation="my_validation2">J K. Rowling</cell>
			<cell style="myInt" validation="my_validation10">2005</cell>
			<cell style="myFloat" validation="my_validation11">29.99</cell>
		</row>
		<row>
			<cell validation="my_validation1">XQuery Kick Start</cell>
			<cell validation="my_validation2">Vaidyanathan Nagarajan</cell>
			<cell style="myInt" validation="my_validation12">2003</cell>
			<cell style="myFloat">49.99</cell>
		</row>
		<row>
			<cell validation="my_validation1">Learning XML</cell>
			<cell validation="my_validation2">Erik T. Ray</cell>
			<cell style="myInt">2003</cell>
			<cell style="myFloat">39.95</cell>
		</row>
	</worksheet>
</workbook>
```

### Mash up:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example9.png?raw=true" alt="Fonts"/>

```XML
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

```XML
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

### A Simple Workbook Defined Column Widths:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example11.png?raw=true" alt="A Simple Workbook Defined Column Widths"/>

```XML
<workbook>
	<worksheet name="Books">
		<columns>
			<column index="0" width="2000"/>
			<column index="1" width="4000"/>
			<column index="2" width="6000"/>
			<column index="3" width="8000"/>
		</columns>
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

### A Simple Pivot Table Using an Area Reference:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example12-1.png?raw=true" alt="A Simple Pivot Table Using an Area Reference"/>
<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example12-2.png?raw=true" alt="A Simple Pivot Table Using an Area Reference"/>

```XML
<workbook>
	<styles>
		<style name="myInt">
			<format type="int"/>
		</style>
	</styles>
	<worksheet name="Books" autofilter="true">
		<row>
			<cell>Group1</cell>
			<cell>Group2</cell>
			<cell>Year</cell>
		</row>
		<row>
			<cell>FOO</cell>
			<cell>BAR</cell>
			<cell style="myInt">2005</cell>
		</row>
		<row>
			<cell>FOO</cell>
			<cell>BAR</cell>
			<cell style="myInt">2021</cell>
		</row>
		<row>
			<cell>Cat</cell>
			<cell>Mouse</cell>
			<cell style="myInt">2005</cell>
		</row>
		<row>
			<cell>Cat</cell>
			<cell>Mouse</cell>
			<cell style="myInt">2006</cell>
		</row>
	</worksheet>
	<worksheet name="Summary">
		<pivot location="A1" dataSheet="Books" dataArea="A1:C5">
			<groupby>
				<column index="0"/>
			</groupby>
			<aggregate>
				<column index="2" action="SUM" name="Sum_Year"/>
				<column index="2" action="COUNT" name="Count_Year"/>
			</aggregate>
			<filter>
				<column index="1"/>
			</filter>
		</pivot>
	</worksheet>
</workbook>
```

### A Simple Pivot Table Using a Table Reference:

<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example13-1.png?raw=true" alt="A Simple Pivot Table Using a Table Reference"/>
<img src="https://github.com/jonbowring/Xml2Xlsx/blob/main/examples/example13-2.png?raw=true" alt="A Simple Pivot Table Using a Table Reference"/>

```XML
<workbook>
	<styles>
		<style name="myInt">
			<format type="int"/>
		</style>
	</styles>
	<worksheet name="Books">
		<table name="My_Table" colStripes="false" rowStripes="true" style="TableStyleMedium3"/>
		<row>
			<cell>Group1</cell>
			<cell>Group2</cell>
			<cell>Year</cell>
		</row>
		<row>
			<cell>FOO</cell>
			<cell>BAR</cell>
			<cell style="myInt">2005</cell>
		</row>
		<row>
			<cell>FOO</cell>
			<cell>BAR</cell>
			<cell style="myInt">2021</cell>
		</row>
		<row>
			<cell>Cat</cell>
			<cell>Mouse</cell>
			<cell style="myInt">2005</cell>
		</row>
		<row>
			<cell>Cat</cell>
			<cell>Mouse</cell>
			<cell style="myInt">2006</cell>
		</row>
	</worksheet>
	<worksheet name="Summary">
		<pivot location="A1" dataSheet="Books" dataTable="My_Table">
			<groupby>
				<column index="0"/>
			</groupby>
			<aggregate>
				<column index="2" action="SUM" name="Sum_Year"/>
				<column index="2" action="COUNT" name="Count_Year"/>
			</aggregate>
			<filter>
				<column index="1"/>
			</filter>
		</pivot>
	</worksheet>
</workbook>
```