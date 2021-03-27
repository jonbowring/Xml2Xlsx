# Xml2Xlsx

## Description:
Xml2Xlsx is an executable jar that enables you to create Excel XLSX files using a simple XML markup. It provides the ability for managing multiple worksheets, formatting, styling, data validation and tables.

## Usage:
```
java -jar Xml2Xlsx-[Version].jar --src [Source XML file] --tgt [Target Excel file]
```

### Example Usage:
```
java -jar xml2xlsx-1.2.2.jar --src "books.xml" --tgt "books.xlsx"
```

## Command Line Options:
Option|Description
------|-----------
--src|Used to specify the location of the input XML file.
--tgt|Used to specify the location of the output Excel file.

## XML Markup:
TBC
XPath|Description
-----|-----------
/workbook/styles/style/@name|Mandatory. The name for the style. This name is used by cells to reference the style to be applied. The name can only contain numbers, letters and underscores.
/workbook/styles/style/align|Optional. Used to define the horizontal and vertical alignment for a style.
/workbook/styles/style/align/@vertical|Optional. Used to define the vertical alignment for an align property. Possible values include "top", "center" and "bottom".
/workbook/styles/style/align/@horizontal|Optional. Used to define the horizontal alignment for an align property. Possible values include "left", "center" and "right".
/workbook/styles/style/border|Optional, can have up to 4 borders defined. Used to define the border style for a cell.
/workbook/styles/style/border/@pos|Mandatory. Used to define which side of the cell the border will be applied to. Possible values include "top", "right", "bottom" and "left".
/workbook/styles/style/border/@type|Optional. Used to define the line style of the border. Possible values include "dash-dot", "dash-dot-dot", "dashed", "dotted", "double", "hair", "medium", "medium-dash-dot", "medium-dash-dot-dot", "medium-dashed", "none", "slanted-dash-dot", "thick" and "thin". If not defined the type "thin" will be applied.
/workbook/styles/style/border/@colour|Optional. Used to define the colour of the border. The colour can be defined as either an rgb colour using the format "rgb(<red>,<green>,<blue>)" (for example "rgb(125,36,210)") or using a pre-defined colour label. Possible pre-defined colour labels include "aqua", "automatic", "black", "black1", "blue", "blue1", "blue-grey", "bright-green", "bright-green1", "brown", "coral", "cornflower-blue", "dark-blue", "dark-green", "dark-red", "dark-teal", "dark-yellow", "gold", "green", "grey-25-percent", "grey-40-percent", "grey-50-percent", "grey-80-percent", "indigo", "lavender", "lemon-chiffon", "light-blue", "light-cornflower-blue", "light-green", "light-orange", "light-turquoise", "light-turquoise1", "light-yellow", "lime", "maroon", "olive-green", "orange", "orchid", "pale-blue", "pink", "pink1", "plum", "red", "red1", "rose", "royal-blue", "sea-green", "sky-blue", "tan", "tan", "turquoise", "turquoise1", "violet", "white", "white1", "yellow" and "yellow1". 


## Examples:
TBC