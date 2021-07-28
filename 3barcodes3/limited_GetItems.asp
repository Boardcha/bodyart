<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetItems__MMColParam
rsGetItems__MMColParam = "1"
If (Request.Form("Bin") <> "") Then 
  rsGetItems__MMColParam = Request.Form("Bin")
  
set SetScanTime = Server.CreateObject("ADODB.Command")
SetScanTime.ActiveConnection = MM_bodyartforms_sql_STRING
SetScanTime.CommandText = "UPDATE TBL_BinNumbers SET BinCountDate = '" & now() & "' WHERE BinNumberID = " & Request.Form("Bin") 
SetScanTime.Execute()

if request.form("ResetBin") = "yes" then 

	set SetScanTime = Server.CreateObject("ADODB.Command")
	SetScanTime.ActiveConnection = MM_bodyartforms_sql_STRING
	SetScanTime.CommandText = "UPDATE QRY_InventoryCount_Limited SET Inventory_TimesScanned = 0 WHERE BinNumber_Detail = " & Request.Form("Bin") 
	SetScanTime.Execute()

end if
  
End If
%>
<%

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
	objCmd.CommandText = "SELECT TOP (100) PERCENT dbo.ProductDetails.ProductDetailID, dbo.ProductDetails.BinNumber_Detail, dbo.jewelry.title, dbo.jewelry.picture, dbo.ProductDetails.Gauge + N' ' + dbo.ProductDetails.Length + N' ' + dbo.ProductDetails.ProductDetail1 AS ProductDescription, dbo.ProductDetails.qty, dbo.ProductDetails.active AS ActiveDetail, dbo.jewelry.active AS ActiveMain, dbo.ProductDetails.Date_InventoryCount, dbo.ProductDetails.Inventory_TimesScanned, dbo.ProductDetails.DateLastPurchased, dbo.jewelry.type FROM dbo.jewelry INNER JOIN dbo.ProductDetails ON dbo.jewelry.ProductID = dbo.ProductDetails.ProductID WHERE BinNumber_Detail = ? AND (dbo.ProductDetails.active = 1) AND (dbo.jewelry.active = 1) ORDER BY dbo.ProductDetails.DateLastPurchased DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("bin",3,1,10,Request.Form("Bin")))
	Set rsGetItems = objCmd.Execute()
%>
<!DOCTYPE html>
<head>
<title>Inventory</title>
<link href="../includes/nav.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.ShowItems {
	display: inline;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	margin: 0px;
}
-->
</style>
</head>

<body onload="document.FRM_ScanItem.Item.focus();" class="materialText">
<form id="FRM_ScanItem" name="FRM_ScanItem" method="post" action="limited_GetItems.asp">
    Scan label: 
  <input name="Item" type="text" id="Item" size="6" />
  &nbsp;&nbsp;
  Qty:
  <input name="qty" type="text" id="qty" size="2" />
  <input name="Bin" type="hidden" id="Bin" value="<%= Request.Form("Bin") %>" /> 
  <input type="submit" name="button" id="button" value="&gt;" />
  <br />
  <strong>BIN # <%= Request.Form("Bin") %>  </strong>
</form>
<br/>




           <% If Not rsGetItems.EOF Or Not rsGetItems.BOF Then %>

<% if request.form("Item") <> "" then ' pull a photo if an item has been scanned %>
<%
Set GetPicture_cmd = Server.CreateObject ("ADODB.Command")
GetPicture_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
GetPicture_cmd.CommandText = "SELECT picture FROM dbo.QRY_InventoryCount_Limited WHERE ProductDetailID = ?" 
GetPicture_cmd.Prepared = true
GetPicture_cmd.Parameters.Append GetPicture_cmd.CreateParameter("param2", 5, 1, -1, Request.Form("Item")) ' adDouble

Set GetPicture = GetPicture_cmd.Execute
%>
 <% If Not GetPicture.EOF Or Not GetPicture.BOF Then %>
           <img src="http://bodyartforms-products.bodyartforms.com/<%= GetPicture.Fields.Item("picture").Value %>" width="90" height="90" />
           <br /><br />

<% end if
end if ' if item scanned has a value pull a picture

Set GetPicture = Nothing
%>
<%
' check to see if item scanned matches a product code in order
 if request.form("Item") <> "" then
 
NotFound = 0
%>

<table>
<thead>
	<th>Qty</th>
	<th>ID</th>
	<th>Name</th>
	<th>Sold</th>
	<th>Active</th>
</thead>
<tbody>
<tr>

<%
While NOT rsGetItems.EOF


if (rsGetItems.Fields.Item("ProductDetailID").Value) = Clng(Request.Form("Item")) then 
NotFound = 1

else 

end if 

  rsGetItems.MoveNext()

Wend
If NotFound = 0 then
response.write "<font size=3 color=red><b>Item #"+ request.form("Item") +" not in bin</b></font><br/>"
End if

rsGetItems.Requery()




' do only if it finds item match
While NOT rsGetItems.EOF

if (rsGetItems.Fields.Item("ProductDetailID").Value) = Clng(Request.Form("Item")) then 


	set SetItemQty = Server.CreateObject("ADODB.Command")
	SetItemQty.ActiveConnection = MM_bodyartforms_sql_STRING
	
			If Request.Form("qty") <> "" then
			
				SetItemQty.CommandText = "UPDATE ProductDetails SET Inventory_TimesScanned = Inventory_TimesScanned + " + Request.Form("qty") + ", Date_InventoryCount = '" & now() & "' WHERE ProductDetailID = " & rsGetItems.Fields.Item("ProductDetailID").Value 

			else

 
				SetItemQty.CommandText = "UPDATE ProductDetails SET Inventory_TimesScanned = Inventory_TimesScanned + 1, Date_InventoryCount = '" & now() & "' WHERE ProductDetailID = " & rsGetItems.Fields.Item("ProductDetailID").Value 
	
	
end if ' compare qty field
	SetItemQty.Execute()

end if ' compare field for match

rsGetItems.MoveNext()
Wend
rsGetItems.Requery()

end if ' if the form field item is empty %>

           <div class="ShowItems">       
  <% 
DetectComplete = 0
While NOT rsGetItems.EOF
%>
<%
Dim rsGetRecentlySold__MMColParam
rsGetRecentlySold__MMColParam = "DetailID"
If (rsGetItems.Fields.Item("ProductDetailID").Value <> "") Then 
  rsGetRecentlySold__MMColParam = rsGetItems.Fields.Item("ProductDetailID").Value
End If
%>
<%
Dim rsGetRecentlySold
Dim rsGetRecentlySold_cmd
Dim rsGetRecentlySold_numRows

Set rsGetRecentlySold_cmd = Server.CreateObject ("ADODB.Command")
rsGetRecentlySold_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecentlySold_cmd.CommandText = "SELECT * FROM dbo.QRY_InventoryCount_GetQTYSold WHERE DetailID = ?" 
rsGetRecentlySold_cmd.Prepared = true
rsGetRecentlySold_cmd.Parameters.Append rsGetRecentlySold_cmd.CreateParameter("param1", 5, 1, -1, rsGetRecentlySold__MMColParam) ' adDouble

Set rsGetRecentlySold = rsGetRecentlySold_cmd.Execute
rsGetRecentlySold_numRows = 0
%> 
       <% if rsGetItems.Fields.Item("Inventory_TimesScanned").Value <> rsGetItems.Fields.Item("qty").Value then ' only show if there are items left to be scanned 
		
DetectComplete = 1 %>


  <% if (rsGetItems.Fields.Item("qty").Value) - (rsGetItems.Fields.Item("Inventory_TimesScanned").Value) < 0 then %>
  <strong><font color="#0033FF">OVER <%=((rsGetItems.Fields.Item("qty").Value) - (rsGetItems.Fields.Item("Inventory_TimesScanned").Value)) * -1%></font></strong>
  <% else %><%=(rsGetItems.Fields.Item("qty").Value) - (rsGetItems.Fields.Item("Inventory_TimesScanned").Value)%><% end if %>
&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;<strong><%=(rsGetItems.Fields.Item("ProductDetailID").Value)%></strong>&nbsp;&nbsp;&nbsp;<%=(rsGetItems.Fields.Item("title").Value)%>&nbsp;<%=(rsGetItems.Fields.Item("ProductDescription").Value)%>&nbsp;&nbsp;
<% If rsGetItems.Fields.Item("DateLastPurchased").Value <> "" AND (rsGetItems.Fields.Item("DateLastPurchased").Value = date()-1 OR rsGetItems.Fields.Item("DateLastPurchased").Value = date()) Then %>
  <strong><font color="#CC0000">SOLD RECENTLY</font></strong>
  <% End If ' end Not rsGetRecentlySold.EOF Or NOT rsGetRecentlySold.BOF %><br /><% end if %> 
<%
rsGetRecentlySold.Close()
Set rsGetRecentlySold = Nothing

  rsGetItems.MoveNext()
  %>
</tr>  
  <%
Wend 
%>
</tbody>
</table>
</div>      
<% if DetectComplete = 0 then %>
<% response.redirect "limited_ScanBin.asp?Complete=Yes" %>
<% end if %>
          <strong><font color="#000000">&nbsp;&nbsp;</font></strong><br />
        &nbsp;<br />
<% End If ' end Not rsGetItems.EOF Or NOT rsGetItems.BOF %>
      </font>
      <% If rsGetItems.EOF And rsGetItems.BOF Then %>
      No bin # found
          <% End If ' end rsGetItems.EOF And rsGetItems.BOF %>
</body>
</html>
<%
rsGetItems.Close()
Set rsGetItems = Nothing
%>
