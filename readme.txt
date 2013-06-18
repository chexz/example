Usage : 
java -jar ElementClassify.jar 
	[--element-table elementTable.xls] 
	[--result result.xls] 
	[--column-mapping columnMapping.xls] 
	source01.xls source02.xls ...

elementTable is used to store the rules to identify element.
eg:
name	xmin	xmax	ymin	ymax	
K		1.1		1.5		2		3.3	
Na		2.1		3.1		4.5		6.8

columnMapping is used to mapping column name used by this app and the experimental tools
eg:
app-column-name		tool-column-name
x					x
y					y
value				value
	
