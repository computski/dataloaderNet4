<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Analytics_dataLoad.aspx.vb" Inherits="PCManalytics.Analytics_dataLoad" Trace="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <title>Data load</title>
     <link rel="stylesheet" type="text/css" href="Styles.css"/>
	
 <script src="javascript\prototype.js" type="text/javascript"></script>
 <script src="javascript\behaviour.js" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
             	<div id="bannerApp"><span style="FONT-VARIANT: normal; FONT-SIZE: 11px; FONT-WEIGHT: normal" id="spanLogin" runat="server">login credentials</span> &nbsp;&nbsp;&nbsp;&nbsp;Pcm Analytics - load data v1&nbsp;</div>
							
    <div id="statusBar" runat="server" EnableViewState="false"></div>
	
    <!--dynamic section-->
   <asp:Menu id="Menu1" Orientation="Horizontal" StaticMenuItemStyle-CssClass="tab" StaticSelectedStyle-CssClass="selectedTab" RenderingMode="table" CssClass="tabs" Runat="server">
        <Items>
        <asp:MenuItem Text="Load data" Value="0" Selected="true" />
        <asp:MenuItem Text="Meta data" Value="1" />
           </Items>    
    </asp:Menu>

 <div class="tabContents">
    <asp:MultiView  id="MultiView"  ActiveViewIndex="0"  Runat="server">
     <asp:View ID="viewLoad" runat="server">
    Load one or more data files here.  Can be in .csv or .zip format
    <br />
    <br />
    <asp:FileUpload ID="filData" multiple="multiple" runat="server" /> <asp:Button ID="btnUpload" Text="Upload file(s)" runat="server" />
    IMPORTANT: max upload size of all files is 10M, you are advised to zip all csv files into a single archive<br />
    <asp:Literal ID="litDataResult" runat="server" />
    <br />
         <asp:button ID="bTest" runat="server" Text="test" Visible="false" />
         <asp:Button ID="Button1" runat="server" Text="Button" Visible="false" />
         <asp:GridView ID="gvDebug" runat="server" />
    
    
    </asp:View>
    <asp:View ID="viewMeta" runat="server">
    Meta data view
    </asp:View>
    
    
    </asp:MultiView>
    </div>


                                 
        </div>
    </form>
</body>
</html>
