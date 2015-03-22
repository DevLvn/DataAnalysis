using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using Microsoft.AnalysisServices;
using System.Xml;
using System.Xml.Linq;
using System.Collections;
using System.Drawing;
using System.Web.UI.DataVisualization.Charting;


namespace OLAPProjectDBMS
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        public static ArrayList catalogList, cubeList, rowList;
        public static string serverName,schema,cube;
        public static List<KeyValuePair<string, ArrayList>> dimensionHierarchyList, measureHierarchyList;
        public static List<ArrayList> axesList, cellDataList;
        public static double maxCellValue;
        
        protected void Page_Load(object sender, EventArgs e)
        {
            ServerName_TextBox.Text = "WIN-NVQQCFCV2JO";
            MDXQuery_TextBox.Text = "SELECT NON EMPTY { [Measures].[Max Order Qty], [Measures].[List Price],[Measures].[Received Qty]}ON COLUMNS,NON EMPTY TOPCOUNT({ ([Product].[Name].[Name].ALLMEMBERS ) },10) ON ROWS FROM [Adventure Works2012]";

            Chart1.Visible = false;
            CheckBox1.Visible = false;
            CheckBox1.AutoPostBack = true;
            
        }

        //*****GENERATES LIST OF SCHEMA BASED ON INPUT SERVERNAME********//
        protected void GenerateSchemaList(object sender, EventArgs e)
        {
            Server server = new Server();
            serverName = ServerName_TextBox.Text;
            try
            {
                server.Connect("Data Source=" + serverName);
                XDocument doc = WriteDiscoverSchemaQuery(server, "DBSCHEMA_CATALOGS");
                ParseDiscoverSchemaXMLA(doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("\nException " + ex.ToString());
            }
            finally
            {
                server.Disconnect();
            }

        }

        protected XDocument WriteDiscoverSchemaQuery(Server server, string type)
        {
            System.Xml.XmlWriter xmlWriter = server.StartXmlaRequest(XmlaRequestType.Undefined);
            xmlWriter.WriteStartElement("Envelope", "http://schemas.xmlsoap.org/soap/envelope/");
            xmlWriter.WriteStartElement("Body");
            xmlWriter.WriteStartElement("Discover", "urn:schemas-microsoft-com:xml-analysis");
            xmlWriter.WriteElementString("RequestType", type);
            xmlWriter.WriteStartElement("Restrictions");
            xmlWriter.WriteEndElement(); // </Restrictions>
            xmlWriter.WriteStartElement("Properties");
            xmlWriter.WriteEndElement(); // </Properties>
            xmlWriter.WriteEndElement(); // </Discover>
            xmlWriter.WriteEndElement(); // </Body>
            xmlWriter.WriteEndElement(); // </Envelope>
            System.Xml.XmlReader xmlReader = server.EndXmlaRequest();
            xmlReader.MoveToContent();
            XDocument doc = XDocument.Load(xmlReader);
            xmlReader.Close();
            return doc;
        }

        protected void ParseDiscoverSchemaXMLA(XDocument doc)
        {
            XNamespace xmlnsa = "http://schemas.xmlsoap.org/soap/envelope/";
            XNode xmlBody = doc.Descendants(xmlnsa + "Body").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());

            xmlnsa = "urn:schemas-microsoft-com:xml-analysis";
            xmlBody = doc.Descendants(xmlnsa + "DiscoverResponse").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());
            xmlnsa = "urn:schemas-microsoft-com:xml-analysis:rowset";

            XElement root = doc.Root;
            IEnumerable<XElement> elements = root.Elements();
            //ArrayList catalogList = new ArrayList();
            catalogList = new ArrayList();
            var list = from item in elements.Descendants(xmlnsa + "CATALOG_NAME") select item;
            SchemaDropDownList.Items.Clear();
            SchemaDropDownList.Items.Add("Choose");
            foreach (var entry in list)
            {
                catalogList.Add(entry.Value.ToString());
                //System.Diagnostics.Debug.WriteLine("\t"+entry.Value);
                SchemaDropDownList.Items.Add(new ListItem(entry.Value.ToString()));
            }

        }

        protected void SchemaDropDownList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SchemaDropDownList.SelectedIndex == 0)
            {
                // System.Diagnostics.Debug.WriteLine("Selected Index 0");
                return;
            }
            schema = SchemaDropDownList.SelectedItem.ToString();
            //System.Diagnostics.Debug.WriteLine("\nItem Selected "+SchemaDropDownList.SelectedItem);
            GenerateCubeList();

        }

        //*****GENERATES LIST OF CUBES BASED ON SCHEMA SELECTED********//
        protected void GenerateCubeList()
        {
            Server server = new Server();
            try
            {
                server.Connect("Data Source=" + serverName);
                XDocument doc = WriteDiscoverCubeQuery(server, "MDSCHEMA_CUBES");
                //System.Diagnostics.Debug.WriteLine("***\n\n" + doc);
                ParseDiscoverCubeXMLA(doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("\nException " + ex.ToString());
            }
            finally
            {
                server.Disconnect();
            }
        }

        protected XDocument WriteDiscoverCubeQuery(Server server, string type)
        {
            System.Xml.XmlWriter xmlWriter = server.StartXmlaRequest(XmlaRequestType.Undefined);
            xmlWriter.WriteStartElement("Envelope", "http://schemas.xmlsoap.org/soap/envelope/");
            xmlWriter.WriteStartElement("Body");
            xmlWriter.WriteStartElement("Discover", "urn:schemas-microsoft-com:xml-analysis");
            xmlWriter.WriteElementString("RequestType", type);
            xmlWriter.WriteStartElement("Restrictions");
            xmlWriter.WriteEndElement(); // </Restrictions>
            xmlWriter.WriteStartElement("Properties");
            xmlWriter.WriteStartElement("PropertyList");
            xmlWriter.WriteElementString("Catalog", schema);
            xmlWriter.WriteEndElement(); // </PropertyList>
            xmlWriter.WriteEndElement(); // </Properties>
            xmlWriter.WriteEndElement(); // </Discover>
            xmlWriter.WriteEndElement(); // </Body>
            xmlWriter.WriteEndElement(); // </Envelope>
            System.Xml.XmlReader xmlReader = server.EndXmlaRequest();
            xmlReader.MoveToContent();
            XDocument doc = XDocument.Load(xmlReader);
            xmlReader.Close();
            return doc;
        }

        protected void ParseDiscoverCubeXMLA(XDocument doc)
        {
            XNamespace xmlnsa = "http://schemas.xmlsoap.org/soap/envelope/";
            XNode xmlBody = doc.Descendants(xmlnsa + "Body").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());

            xmlnsa = "urn:schemas-microsoft-com:xml-analysis";
            xmlBody = doc.Descendants(xmlnsa + "DiscoverResponse").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());
            xmlnsa = "urn:schemas-microsoft-com:xml-analysis:rowset";

            XElement root = doc.Root;
            IEnumerable<XElement> elements = root.Elements();
            //ArrayList catalogList = new ArrayList();
            cubeList = new ArrayList();
            var list = from item in elements.Descendants(xmlnsa + "CUBE_NAME") select item;
            CubeDropDownList.Items.Clear();
            CubeDropDownList.Items.Add("Choose");
            foreach (var entry in list)
            {
                catalogList.Add(entry.Value.ToString());
                // System.Diagnostics.Debug.WriteLine("\t"+entry.Value);
                CubeDropDownList.Items.Add(new ListItem(entry.Value.ToString()));
            }

        }

        protected void CubeDropDownList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CubeDropDownList.SelectedIndex == 0)
            {
                //System.Diagnostics.Debug.WriteLine("Selected Index 0");
                return;
            }
            cube = CubeDropDownList.SelectedItem.ToString();
            //System.Diagnostics.Debug.WriteLine("\nItem Selected " + CubeDropDownList.SelectedItem);
            GeneraterTreeView();
        }

        //*****GENERATES TREEVIEW WITH ALL DIMENSIONS AND MEASURES IN CUBE SELECTED********//
        protected void GeneraterTreeView()
        {
            //List<KeyValuePair<string, ArrayList>> dimensionHierarchyList = new List<KeyValuePair<string, ArrayList>>();

            string measureDVM, dimensionDVM;
            measureDVM = "SELECT [MEASUREGROUP_NAME] AS [FOLDER],[MEASURE_CAPTION] AS [MEASURE] FROM $SYSTEM.MDSCHEMA_MEASURES WHERE CUBE_NAME  ='" + cube + "' ORDER BY [MEASUREGROUP_NAME]";
            dimensionDVM = "SELECT [DIMENSION_UNIQUE_NAME] AS [DIMENSION], HIERARCHY_CAPTION AS [DIMENSION_ATTRIBUTE] FROM $system.MDSchema_hierarchies WHERE CUBE_NAME  ='" + cube + "' ORDER BY [DIMENSION_UNIQUE_NAME]";
            Server server = new Server();

            try
            {
                server.Connect("Data Source=" + serverName);
                XDocument measureDoc = ExecuteDVMQuery(server, measureDVM);
                measureHierarchyList = ParseExecuteDVMXMLA(measureDoc, "FOLDER", "MEASURE");
                XDocument dimensionDoc = ExecuteDVMQuery(server, dimensionDVM);
                dimensionHierarchyList = ParseExecuteDVMXMLA(dimensionDoc, "DIMENSION", "DIMENSION_ATTRIBUTE");
                dimensionHierarchyList.RemoveAt(0);

                TreeNode tempNode = new TreeNode(cube);
                tempNode.SelectAction = TreeNodeSelectAction.None;
                TreeView1.Nodes.Add(tempNode);
                tempNode = new TreeNode("Measures");
                tempNode.SelectAction = TreeNodeSelectAction.None;
                TreeView1.Nodes[0].ChildNodes.Add(tempNode);
                tempNode = new TreeNode("Dimensions");
                tempNode.SelectAction = TreeNodeSelectAction.None;
                TreeView1.Nodes[0].ChildNodes.Add(tempNode);

                int dimCount = 0;
                int mesCount = 0;
                foreach (var entry in measureHierarchyList)
                {
                    tempNode = new TreeNode(entry.Key.ToString());
                    tempNode.SelectAction = TreeNodeSelectAction.None;
                    TreeView1.Nodes[0].ChildNodes[0].ChildNodes.Add(tempNode);
                    foreach (var mes in entry.Value)
                    {
                        tempNode = new TreeNode(mes.ToString());
                        tempNode.SelectAction = TreeNodeSelectAction.None;
                        TreeView1.Nodes[0].ChildNodes[0].ChildNodes[dimCount].ChildNodes.Add(tempNode);
                    }
                    dimCount++;
                }

                foreach (var entry in dimensionHierarchyList)
                {
                    tempNode = new TreeNode(entry.Key.ToString());
                    tempNode.SelectAction = TreeNodeSelectAction.None;
                    TreeView1.Nodes[0].ChildNodes[1].ChildNodes.Add(tempNode);
                    foreach (var mes in entry.Value)
                    {
                        tempNode = new TreeNode(mes.ToString());
                        tempNode.SelectAction = TreeNodeSelectAction.None;
                        TreeView1.Nodes[0].ChildNodes[1].ChildNodes[mesCount].ChildNodes.Add(tempNode);
                    }
                    mesCount++;
                }

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("\nException " + ex.ToString());
            }
            finally
            {
                server.Disconnect();
            }

        }

        protected XDocument ExecuteDVMQuery(Server server, string query)
        {
            System.Xml.XmlWriter xmlWriter = server.StartXmlaRequest(XmlaRequestType.Undefined);
            WriteDVMEnvelope(xmlWriter, query);
            System.Xml.XmlReader xmlReader = server.EndXmlaRequest();
            xmlReader.MoveToContent();
            XDocument doc = XDocument.Load(xmlReader);
            xmlReader.Close();

            return doc;
        }

        private static void WriteDVMEnvelope(System.Xml.XmlWriter xmlWriter, string query)
        {

            xmlWriter.WriteStartElement("Envelope", "http://schemas.xmlsoap.org/soap/envelope/");
            xmlWriter.WriteStartElement("Body");
            xmlWriter.WriteStartElement("Execute", "urn:schemas-microsoft-com:xml-analysis");
            xmlWriter.WriteStartElement("Command");
            xmlWriter.WriteElementString("Statement", query);
            xmlWriter.WriteEndElement(); // </Command>
            xmlWriter.WriteStartElement("Properties");
            xmlWriter.WriteEndElement(); // </Properties>
            xmlWriter.WriteEndElement(); // </Execute>
            xmlWriter.WriteEndElement(); // </Body>
            xmlWriter.WriteEndElement(); // </Envelope>
        }

        protected List<KeyValuePair<string, ArrayList>> ParseExecuteDVMXMLA(XDocument doc, string var1, string var2)
        {
            List<KeyValuePair<string, ArrayList>> keyValueList = new List<KeyValuePair<string, ArrayList>>();
            XNamespace xmlnsa = "http://schemas.xmlsoap.org/soap/envelope/";
            XNode xmlBody = doc.Descendants(xmlnsa + "Body").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());

            xmlnsa = "urn:schemas-microsoft-com:xml-analysis";
            xmlBody = doc.Descendants(xmlnsa + "ExecuteResponse").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());
            xmlnsa = "urn:schemas-microsoft-com:xml-analysis:rowset";

            XElement root = doc.Root;
            IEnumerable<XElement> elements = root.Elements();
            //ArrayList catalogList = new ArrayList();

            IEnumerable<XElement> folderlist = from item in elements.Descendants(xmlnsa + var1) select item;
            IEnumerable<XElement> measurelist = from item in elements.Descendants(xmlnsa + var2) select item;
            var enumerator1 = folderlist.GetEnumerator();
            var enumerator2 = measurelist.GetEnumerator();

            XElement group, mValue;
            string groupName, curValue, prevGroup;
            prevGroup = "";
            ArrayList valueList = new ArrayList();
            while (enumerator1.MoveNext() && enumerator2.MoveNext())
            {
                group = enumerator1.Current;
                mValue = enumerator2.Current;
                groupName = group.Value.ToString();
                groupName = groupName.Replace("[", "");
                groupName = groupName.Replace("]", "");
                curValue = mValue.Value.ToString();
                if (prevGroup != groupName && prevGroup != "")
                {
                    valueList.Sort();
                    keyValueList.Add(new KeyValuePair<string, ArrayList>(prevGroup, valueList));
                    valueList = new ArrayList();
                }
                valueList.Add(curValue);
                prevGroup = groupName;
            }
            valueList.Sort();
            keyValueList.Add(new KeyValuePair<string, ArrayList>(prevGroup, valueList));

            //foreach (var entry in keyValueList)
            //{
            //    System.Diagnostics.Debug.WriteLine("\n\tKey: " + entry.Key);
            //    foreach (var mes in entry.Value)
            //        System.Diagnostics.Debug.WriteLine("\t" + mes);
            //}

            return keyValueList;
        }
        
        //***EXECUTE MDX QUERY*********//
        protected void ExecuteMDXQuery(object sender, EventArgs e)
        {
            string query = MDXQuery_TextBox.Text;
            Server server = new Server();
            try
            {
                server.Connect("Data Source=" + serverName);
                XDocument doc = StartAndEndXmlaRequest(server, query);
                ParseXMLA(doc);
                GenerateChart();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("\nException " + ex.ToString());
            }
            finally
            {
                server.Disconnect();
                CheckBox1.Visible = true;
            }
        }

        private static XDocument StartAndEndXmlaRequest(Server server, string query)
        {
            // Step 1: start the XML/A request.
            System.Xml.XmlWriter xmlWriter = server.StartXmlaRequest(XmlaRequestType.Undefined);

            // Step 2: write the XML/A request.
            WriteSoapEnvelope(xmlWriter, query);

            // Step 3: end the XML/A request and get the System.Xml.XmlReader for parsing the result from server.
            System.Xml.XmlReader xmlReader = server.EndXmlaRequest();

            // Step 4: read/parse the XML/A response from server.
            xmlReader.MoveToContent();
            XDocument doc = XDocument.Load(xmlReader);

            // Step 5: close the System.Xml.XmlReader, to release the connection for future use.
            xmlReader.Close();

            return doc;
        }

        private static void WriteSoapEnvelope(System.Xml.XmlWriter xmlWriter, string query)
        {
            xmlWriter.WriteStartElement("Envelope", "http://schemas.xmlsoap.org/soap/envelope/");
            xmlWriter.WriteStartElement("Body");
            xmlWriter.WriteStartElement("Execute", "urn:schemas-microsoft-com:xml-analysis");
            xmlWriter.WriteStartElement("Command");
            xmlWriter.WriteElementString("Statement", query);
            xmlWriter.WriteEndElement(); // </Command>
            xmlWriter.WriteStartElement("Properties");
            xmlWriter.WriteStartElement("PropertyList");
            xmlWriter.WriteElementString("Catalog", schema);
            xmlWriter.WriteElementString("Format", "Multidimensional");
            xmlWriter.WriteElementString("Content", "Data");
            xmlWriter.WriteEndElement(); // </PropertyList>
            xmlWriter.WriteEndElement(); // </Properties>
            xmlWriter.WriteEndElement(); // </Execute>
            xmlWriter.WriteEndElement(); // </Body>
            xmlWriter.WriteEndElement(); // </Envelope>
        }

        private static void ParseXMLA(XDocument doc)
        {
            XNamespace xmlnsa = "http://schemas.xmlsoap.org/soap/envelope/";
            XNode xmlBody = doc.Descendants(xmlnsa + "Body").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());
            
            xmlnsa = "urn:schemas-microsoft-com:xml-analysis";
            xmlBody = doc.Descendants(xmlnsa + "ExecuteResponse").First().FirstNode;
            doc = XDocument.Parse(xmlBody.ToString());
            xmlnsa = "urn:schemas-microsoft-com:xml-analysis:mddataset";
            
            //*******SAVING METADATA ABOUT AXES AS XDOCUMENT
            xmlBody = doc.Descendants(xmlnsa + "root").First().FirstNode.NextNode;
            //System.Diagnostics.Debug.WriteLine("**Metadata***\n" + doc);
            XDocument metaDataDoc = XDocument.Parse(xmlBody.ToString());
            //System.Diagnostics.Debug.WriteLine("**Metadata***\n"+metaDataDoc);
            
            //*******SAVING CELLDATA AS XDOCUMENT
            xmlBody = doc.Descendants(xmlnsa + "root").First().FirstNode.NextNode.NextNode;
            XDocument cellDataDoc = XDocument.Parse(xmlBody.ToString());
            //System.Diagnostics.Debug.WriteLine("\n\n**celldata***\n" + cellDataDoc);

            //*******EXTARCT NUMBER & NAME OF COLUMNS AND ROWS
            axesList = new List<ArrayList>();
            Tuple<List<ArrayList>,int> tuple = ExtractAxesDetails(metaDataDoc, xmlnsa);
            axesList = tuple.Item1;
            int columnCount = tuple.Item2;
            
            //*******EXTARCT CELL DATA FROM XMLA
            cellDataList = new List<ArrayList>();
            cellDataList = ExtracCellDetails(cellDataDoc, xmlnsa, columnCount);
        }

        private static Tuple<List<ArrayList>,int> ExtractAxesDetails(XDocument metaDataDoc,XNamespace xmlnsa)
        {
            XElement root = metaDataDoc.Root;
            IEnumerable<XElement> elements = root.Elements();
            List<ArrayList> tempAxesList = new List<ArrayList>();
            int columnCount = 0;

            string currentAxis;
            foreach (XElement element in elements)
            {
                currentAxis = (element.Attribute("name")).Value.ToString();
                //System.Diagnostics.Debug.WriteLine("**Element\n" + currentAxis);

                if (elements.Count() == 3 && !currentAxis.Equals("SlicerAxis"))
                {

                    if (currentAxis == "Axis0")
                    {
                        ArrayList Axis0 = new ArrayList();
                        var list = from item in element.Descendants(xmlnsa + "UName") select item;
                        foreach (var entry in list)
                        {
                            string temp = entry.Value.ToString();
                            temp = temp.Replace("[", "");
                            temp = temp.Replace("]", "");
                            temp = temp.Replace(".", "_");
                            Axis0.Add(temp);
                            columnCount++;
                        }
                        tempAxesList.Add(Axis0);
                    }
                    else
                    {
                        ArrayList Axis1 = new ArrayList();
                        var list = from item in element.Descendants(xmlnsa + "Member") select item;
                        string temp = list.First().Attribute("Hierarchy").Value.ToString();
                        temp = temp.Replace("[", "");
                        temp = temp.Replace("]", "");
                        temp = temp.Replace(".", "_");
                        Axis1.Add(temp);
                        tempAxesList.Add(Axis1);
                        var tempRowList = from item in element.Descendants(xmlnsa + "Caption") select item;
                        rowList = new ArrayList();
                        foreach (var entry in tempRowList)
                        {
                            rowList.Add(entry.Value);
                            //System.Diagnostics.Debug.WriteLine("\nRowVal\t" + entry.Value);
                        }
                    }
                }
            }

            //foreach (ArrayList axis in tempAxesList)
            //{
            //    System.Diagnostics.Debug.WriteLine("\nAxis:");
            //    foreach (Object a in axis)
            //        System.Diagnostics.Debug.WriteLine("\n" + a.ToString());
            //}

            return Tuple.Create<List<ArrayList>, int>(tempAxesList, columnCount);
        }

        private static List<ArrayList> ExtracCellDetails(XDocument cellDataDoc, XNamespace xmlnsa, int columnCount)
        {
            XElement cellRoot = cellDataDoc.Root;
            IEnumerable<XElement> cellElements = cellRoot.Elements();
            List<ArrayList> tempCellDataList = new List<ArrayList>();  
            var enumerator = cellElements.GetEnumerator();
            string currentOrdinal,currentData;
            int ordinalValue, count;
            double cellValue;
            ArrayList rowData = new ArrayList();
            count = 0;
            //maxValueEachAxis = new ArrayList();
            double maxCellValue = -1;
            while (enumerator.MoveNext())
            {
               // count++;
                XElement element = enumerator.Current;
                currentOrdinal = (element.Attribute("CellOrdinal")).Value.ToString();
                ordinalValue = Convert.ToInt32(currentOrdinal);
                currentData = element.Descendants(xmlnsa + "FmtValue").First().Value;
                cellValue=Convert.ToDouble(currentData);
                if (maxCellValue > cellValue)
                    maxCellValue = cellValue;
               // System.Diagnostics.Debug.WriteLine("\n***count " + count + "\tordinalValue " + ordinalValue + "\t**Value " + currentData);
                
                
                while (count <= ordinalValue)
                {
                    if (count == ordinalValue && ((count % columnCount) == columnCount - 1))
                    {
                        //System.Diagnostics.Debug.WriteLine("Adding " + ordinalValue);
                        rowData.Add(currentData);
                        tempCellDataList.Add(rowData);
                        //maxValueEachAxis.Add(maxCellValue);
                        //maxCellValue = -1;
                        rowData = new ArrayList();
                    }
                    else if (count == ordinalValue)
                    {
                        //System.Diagnostics.Debug.WriteLine("Adding " + ordinalValue);
                        rowData.Add(currentData);
                    }
                    else if ((count % columnCount) == columnCount - 1)
                    {
                        //System.Diagnostics.Debug.WriteLine("Adding null");
                        rowData.Add("null");
                        tempCellDataList.Add(rowData);
                        //maxValueEachAxis.Add(maxCellValue);
                        //maxCellValue = -1;
                        rowData = new ArrayList();
                    }
                    else
                    {
                        //System.Diagnostics.Debug.WriteLine("Adding null");
                        rowData.Add("null");
                    }
                    count++;
                }
                 
            }


            //foreach (ArrayList cell in tempCellDataList)
            //{
            //    System.Diagnostics.Debug.WriteLine("\nCell Value:");
            //    foreach (Object a in cell)
            //    { System.Diagnostics.Debug.WriteLine("\t" + a.ToString()); }
            //}

            return tempCellDataList;
        }

        protected void GenerateChart()
        {
            var axes = axesList.GetEnumerator();
            axes.MoveNext();
            ArrayList columnList = axes.Current;
            int height = 400;
            int width = 90;
            int colCount = columnList.Count;
            int rowCount = rowList.Count;
            if (colCount > 1)
            {
                height += (400 * (colCount - 1));
            }

            if (rowCount>1)
            {
                width+=(90*(rowCount-1));
            }

            Chart1.ChartAreas.Clear();

            //Code for multile Chart Area
            axes.MoveNext();
            ArrayList rowName = axes.Current;
            int count = 1;
            foreach (var column in columnList)
           // for (int i = 1; i <= colCount; i++)
            {
                string chartAreaName="ChartArea"+count;
                System.Diagnostics.Debug.WriteLine("Name "+chartAreaName);
                Chart1.ChartAreas.Add(new ChartArea(chartAreaName));
                Chart1.ChartAreas[chartAreaName].AxisX.Minimum = 0;
                Chart1.ChartAreas[chartAreaName].AxisY.Minimum = 0;
                //*****AXES TITLE & LABEL******/

                Chart1.ChartAreas[chartAreaName].AxisX.Title = rowName[0].ToString();
                Chart1.ChartAreas[chartAreaName].AxisX.TitleFont = new Font(new FontFamily("Calibri"), 10, FontStyle.Bold);
                Chart1.ChartAreas[chartAreaName].AxisX.TitleForeColor = Color.WhiteSmoke;
                Chart1.ChartAreas[chartAreaName].AxisX.LabelStyle.ForeColor = Color.WhiteSmoke;
                Chart1.ChartAreas[chartAreaName].AxisX.LabelStyle.Font = new Font(new FontFamily("Calibri"), 9, FontStyle.Bold);
                Chart1.ChartAreas[chartAreaName].AxisY.Title = column.ToString();
                Chart1.ChartAreas[chartAreaName].AxisY.TitleFont = new Font(new FontFamily("Calibri"), 10, FontStyle.Bold);
                Chart1.ChartAreas[chartAreaName].AxisY.TitleForeColor = Color.WhiteSmoke;
                Chart1.ChartAreas[chartAreaName].AxisY.LabelStyle.ForeColor = Color.WhiteSmoke;
                Chart1.ChartAreas[chartAreaName].AxisY.LabelStyle.Font = new Font(new FontFamily("Calibri"), 9, FontStyle.Bold);
                Chart1.ChartAreas[chartAreaName].AxisX.MajorGrid.LineColor = Color.LightGray;
                Chart1.ChartAreas[chartAreaName].AxisY.MajorGrid.LineColor = Color.LightGray;

                //Chart1.ChartAreas[chartAreaName].BorderDashStyle = ChartDashStyle.Solid;
                //Chart1.ChartAreas[chartAreaName].BorderColor = Color.Black;
                //Chart1.ChartAreas[chartAreaName].BorderWidth = 2;

                count++;
            }


            Chart1.Visible = true;
            Chart1.Series.Clear();
            //Chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;
            //Chart1.ChartAreas["ChartArea1"].Area3DStyle.Perspective = 30;

            Chart1.BackColor = Color.Transparent;
            //Chart1.ChartAreas["ChartArea1"].AxisY.MajorGrid.Interval = Convert.ToInt32(maxCellValue / 20);


            Chart1.Titles.Clear();
            var title = new Title
            {
                Text = "MDX RESULT",
                ForeColor = Color.Wheat,
                Font = new Font(new FontFamily("Calibri"), 14, FontStyle.Bold),
                Alignment = System.Drawing.ContentAlignment.TopLeft,
                ToolTip = "Title"
            };
            Chart1.Titles.Add(title);


            
            Chart1.Height = height;
            Chart1.Width = width;

            string[] color = { "333399", "660033", "003300", "666633", "6600CC", "79405D", "0099CC" };
            string tempvar = "";
            //Color = Color.FromArgb(200, System.Drawing.ColorTranslator.FromHtml("#" + color[count])),
            count = 1;
            foreach (var column in columnList)
            {
                string columnName = column.ToString();
                var series = new Series
                {
                    Name = columnName,
                    Color = Color.FromArgb(200, System.Drawing.ColorTranslator.FromHtml("#" + color[new Random().Next(0,color.Count())])),
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column
                };
                Chart1.Series.Add(series);
                Chart1.Series[columnName].ChartArea = "ChartArea" + count;
                Chart1.Series[columnName].IsVisibleInLegend = true;
                //Chart1.Series[columnName].ChartType = SeriesChartType.Bar;
                Chart1.Series[columnName]["DrawingStyle"] = "Cylinder";

                //****DATAPOINT LABEL******//
                //Chart1.Series[columnName]["LabelStyle"] = "BottomLeft";
                Chart1.Series[columnName].Font = new Font(new FontFamily("Calibri"), 8);
                Chart1.Series[columnName].LabelToolTip = "#VAL";
                Chart1.Series[columnName].IsValueShownAsLabel = true;
                //Chart1.Series[columnName]["BarLabelStyle"] = "Outside";//values Outside, Left, Right, Center 
                Chart1.Series[columnName].SmartLabelStyle.Enabled = true;
                Chart1.Series[columnName].SmartLabelStyle.IsMarkerOverlappingAllowed = false;
                Chart1.Series[columnName].LabelAngle = 0;
                //Chart1.Series[columnName].LabelForeColor = Color.FromArgb(250,Color.Black);

                tempvar = columnName;
                count++;
            }



            //***GENERATING DATAPOINTS****//
            //rowList->AL//L<AL>:cellDataList,axesList
            var row = rowList.GetEnumerator();
            axes = axesList.GetEnumerator();
            axes.MoveNext();
            columnList = axes.Current;

            var data = cellDataList.GetEnumerator();

            foreach (ArrayList lis in cellDataList)
            {
                count = 0;
                row.MoveNext();
                axes.MoveNext();
                foreach (var column in columnList)
                {
                    string columnName = column.ToString();
                    Chart1.Series[columnName].Points.AddXY(row.Current.ToString(), lis[count]);
                    count++;
                }
            }

        }

        
        protected void SingleChart()
        {
            ChartArea ChartArea1 = new ChartArea();
            Chart1.ChartAreas.Add(ChartArea1);
            Chart1.Visible = true;
            Chart1.Series.Clear();
            Chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;

            //Chart1.ChartAreas["ChartArea1"].Area3DStyle.Perspective = 30;
            Chart1.ChartAreas["ChartArea1"].AxisX.Minimum = 0;
            Chart1.ChartAreas["ChartArea1"].AxisY.Minimum = 0;
            Chart1.BackColor = Color.Transparent;
            Chart1.ChartAreas["ChartArea1"].AxisY.MajorGrid.Interval = Convert.ToInt32(maxCellValue / 20);
            

            Chart1.Titles.Clear();
            var title = new Title
            {
                Text = "MDX RESULT",
                ForeColor = Color.Wheat,
                Font = new Font(new FontFamily("Calibri"), 14, FontStyle.Bold),
                Alignment = System.Drawing.ContentAlignment.TopLeft,
                ToolTip = "Title"
            };
            Chart1.Titles.Add(title);

            Chart1.Legends.Add(new Legend("Legend1"));

            var axes = axesList.GetEnumerator();
            axes.MoveNext();
            ArrayList columnList = axes.Current;

            int height = 400;
            int width = 500;
            int colCount = columnList.Count;
            if (colCount > 2)
            {
                height += (100 * (colCount - 2));
                width += (100 * (colCount - 2));
            }
            Chart1.Height = height;
            Chart1.Width = width;
            Chart1.ChartAreas["ChartArea1"].Position.Y = 5;
            Chart1.ChartAreas["ChartArea1"].Position.Width = 80;
            Chart1.ChartAreas["ChartArea1"].Position.Height = 80;
            string[] color = { "333399", "660033", "003300", "666633", "6600CC", "79405D" };
            string tempvar = "";

            int count = 0;
            foreach (var column in columnList)
            {
                string columnName = column.ToString();
                System.Diagnostics.Debug.WriteLine("\n***columnName\t" + columnName);
                var series = new Series
                {
                    Name = columnName,
                    Color = Color.FromArgb(200, System.Drawing.ColorTranslator.FromHtml("#" + color[count])),
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Line
                };

                Chart1.Series.Add(series);
                Chart1.Series[columnName].Legend = "Legend1";
                Chart1.Series[columnName].IsVisibleInLegend = true;
                Chart1.Series[columnName].ChartType = SeriesChartType.Bar;
                Chart1.Series[columnName]["DrawingStyle"] = "Cylinder";

                //****DATAPOINT LABEL******//
                //Chart1.Series[columnName]["LabelStyle"] = "BottomLeft";
                Chart1.Series[columnName].Font = new Font(new FontFamily("Calibri"), 1);
                Chart1.Series[columnName].LabelToolTip = "#VAL";
                Chart1.Series[columnName].IsValueShownAsLabel = true;
                //Chart1.Series[columnName]["BarLabelStyle"] = "Outside";//values Outside, Left, Right, Center 
                //Chart1.Series[columnName].SmartLabelStyle.Enabled = true;
                //Chart1.Series[columnName].SmartLabelStyle.IsMarkerOverlappingAllowed = false;
                //Chart1.Series[columnName].LabelAngle = 0;
                //Chart1.Series[columnName].LabelForeColor = Color.FromArgb(250,Color.Black);

                tempvar = columnName;
                count++;
            }
            Chart1.Legends["Legend1"].Alignment = System.Drawing.StringAlignment.Far;

            //*****AXES TITLE & LABEL******/
            axes.MoveNext();
            columnList = axes.Current;
            Chart1.ChartAreas["ChartArea1"].AxisX.Title = columnList[0].ToString();
            Chart1.ChartAreas["ChartArea1"].AxisX.TitleFont = new Font(new FontFamily("Calibri"), 11, FontStyle.Bold);
            Chart1.ChartAreas["ChartArea1"].AxisX.TitleForeColor = Color.WhiteSmoke;
            Chart1.ChartAreas["ChartArea1"].AxisX.LabelStyle.ForeColor = Color.WhiteSmoke;
            Chart1.ChartAreas["ChartArea1"].AxisX.LabelStyle.Font = new Font(new FontFamily("Calibri"), 9, FontStyle.Bold);
            Chart1.ChartAreas["ChartArea1"].AxisY.LabelStyle.ForeColor = Color.WhiteSmoke;
            Chart1.ChartAreas["ChartArea1"].AxisY.LabelStyle.Font = new Font(new FontFamily("Calibri"), 9, FontStyle.Bold);
            //***GENERATING DATAPOINTS****//
            //rowList->AL//L<AL>:cellDataList,axesList
            var row = rowList.GetEnumerator();
            axes = axesList.GetEnumerator();
            axes.MoveNext();
            columnList = axes.Current;

            var data = cellDataList.GetEnumerator();

            foreach (ArrayList lis in cellDataList)
            {
                count = 0;
                row.MoveNext();
                axes.MoveNext();
                foreach (var column in columnList)
                {
                    string columnName = column.ToString();
                    Chart1.Series[columnName].Points.AddXY(row.Current.ToString(), lis[count]);
                    count++;
                }
            }

        }

        protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("\nChecked\t" + CheckBox1.Checked);
            if (CheckBox1.Checked)
            {
                SingleChart();
            }
            else
            {
                GenerateChart();
            }
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            Page.Response.Redirect("/Default.aspx");
        }
    }
}