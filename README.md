# Xml2Xlsx

## Usage
java -jar Xml2Xlsx.jar --src [Source XML file] --tgt [Target Excel file]

## Example source XML file
```
<workbook>
	<worksheet name="Books">
		<row>
			<cell type="string">Title</cell>
			<cell type="string">Author</cell>
			<cell type="string">Year</cell>
			<cell type="string">Price</cell>
		</row>
		<row>
			<cell type="formula">1+1</cell>
			<cell type="string">Giada De Laurentiis</cell>
			<cell type="date">2005-11-26</cell>
			<cell type="float">30.00</cell>
		</row>
		<row>
			<cell type="string">Harry Potter</cell>
			<cell type="string">J K. Rowling</cell>
			<cell type="int">2005</cell>
			<cell type="float">29.99</cell>
		</row>
		<row>
			<cell type="string">XQuery Kick Start</cell>
			<cell type="string">James McGovern</cell>
			<cell type="int">2003</cell>
			<cell type="float">49.99</cell>
		</row>
		<row>
			<cell type="string">Learning XML</cell>
			<cell type="string">Erik T. Ray</cell>
			<cell type="int">2003</cell>
			<cell type="float">39.95</cell>
		</row>
	</worksheet>
</workbook>
```
