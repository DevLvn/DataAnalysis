# DataAnalysis

##Description
List of schemas and cubes are dynamically generated based on SSAS server name passed. Discover command in XML for Analysis are wrapped in SOAP envelop and sent to SSAS; the XML response is parsed and data is extracted to build a TreeView containing all dimensions and measures in the cube.
User issues an MDX query which is converted to Execute command in XML for Analysis and sent to SSAS, the response is parsed and result is used to populate chart. 

##Prerequisite:
atlas one OLAP cube is built and deployed in Microsoft SQL Server Analysis Services(SSAS)

##Sample:
'Output Sample’ folder contains html file recorded during various phase of execution.
* SSAS Data Analysis 1.htm : Default page when launched
* SSAS Data Analysis 2.htm : TreeView containing measure and dimensions for the selected OLAP cube populated on left pane
* SSAS Data Analysis 3.htm : MDX query is executed, and chart generated for individual measure and dimension combination
* SSAS Data Analysis 4.htm: Single chart combining all the measures is generated for comparison. 



##Summary
Analyzing OLAP cube stored in Microsoft SQL Server Analysis Services(SSAS) from web application using XML for Analysis. 
