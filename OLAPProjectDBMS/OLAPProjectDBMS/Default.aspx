<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="OLAPProjectDBMS.WebForm1" %>

<%@ Register assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" namespace="System.Web.UI.DataVisualization.Charting" tagprefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<link href="StyleSheet1.css" rel="Stylesheet" type="text/css"/>
    <title>SSAS Data Analysis</title>
</head>
<body>
    <form id="form1" runat="server">
    <h1 id="mainHeader">DATA ANALYSIS</h1>
    <div id="myDetails">
        <p>CMPT 740</p>
        <p>Suraiya Hameed</p>
        <p>301246668</p>
    </div>
    <div class="space"></div>
    <div id="treeView">
    <p>        
        <asp:TreeView ID="TreeView1" runat="server" ImageSet="Arrows">
        <LeafNodeStyle CssClass="leafNode" />
        <NodeStyle CssClass="treeNode" />
        <RootNodeStyle CssClass="rootNode" />
        <SelectedNodeStyle CssClass="selectNode" />
        </asp:TreeView>
    </p>
    </div>
    <div id="queryView">
    <p>
        <asp:Label ID="Server_Label" runat="server" Text="Server Name" CssClass="label"></asp:Label>
        <asp:TextBox ID="ServerName_TextBox" runat="server" CssClass="editbox"></asp:TextBox>
        <asp:Button ID="Button2" runat="server" Text="Generate Schema List" onclick="GenerateSchemaList" CssClass="button"/>
    </p>
    <p>
        <asp:Label ID="Schema_Label" runat="server" Text="Schema List" CssClass="label"></asp:Label>    
        <asp:DropDownList ID="SchemaDropDownList" runat="server" AutoPostBack="True" CssClass="editbox"
            onselectedindexchanged="SchemaDropDownList_SelectedIndexChanged">
        </asp:DropDownList>
    </p>
    <p>
        <asp:Label ID="Cube_Label" runat="server" Text="Cube List" CssClass="label"></asp:Label>
        <asp:DropDownList ID="CubeDropDownList" runat="server" AutoPostBack="True" CssClass="editbox"
            onselectedindexchanged="CubeDropDownList_SelectedIndexChanged">
        </asp:DropDownList>
    </p>
    <p>
        <asp:Label ID="MDXQuery_Label" runat="server" Text="MDX Query" CssClass="label"></asp:Label>
        <asp:TextBox ID="MDXQuery_TextBox" runat="server" CssClass="MDXeditbox" TextMode="Multiline" Columns="50" Rows="5"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" Text="Execute MDX" onclick="ExecuteMDXQuery" CssClass="button"/>
    </p>
    <div class="space">
    <p>       
        <asp:Button ID="Button3" runat="server" Text="Reset"  onclick="Button3_Click" CssClass="Resetbutton"/>
    </p>
    </div>
    </div>
    <div id="queryResult">
    <div class="space"></div>
    <div id="resulttable"></div>
    <div class="space"></div>
    <div>
        <asp:CheckBox ID="CheckBox1" runat="server"  
            oncheckedchanged="CheckBox1_CheckedChanged" Text="Single Chart" 
            TextAlign="Left" />
        </div>
    <div class="space"></div>
    <p><asp:Chart ID="Chart1" runat="server"  >
        </asp:Chart>
        </p>
    </div>
    </form>
</body>
</html>
