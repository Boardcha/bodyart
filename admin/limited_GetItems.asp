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
Dim rsGetItems__MMColParam2
rsGetItems__MMColParam2 = "1"
If (Request.Form("Bin") <> "") Then 
  rsGetItems__MMColParam2 = Request.Form("Bin")
End If
%>
<%
Dim rsGetItems
Dim rsGetItems_cmd
Dim rsGetItems_numRows

Set rsGetItems_cmd = Server.CreateObject ("ADODB.Command")
rsGetItems_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetItems_cmd.CommandText = "SELECT * FROM dbo.QRY_InventoryCount_Limited WHERE BinNumber_Detail = ?" 
rsGetItems_cmd.Prepared = true
rsGetItems_cmd.Parameters.Append rsGetItems_cmd.CreateParameter("param1", 5, 1, -1, rsGetItems__MMColParam) ' adDouble

Set rsGetItems = rsGetItems_cmd.Execute
rsGetItems_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetItems_numRows = rsGetItems_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
 %>
<%
'GetPicture_cmd.Close()
'Set GetPicture_cmd = Nothing
%>
<%
' check to see if item scanned matches a product code in order
 if request.form("Item") <> "" then
 
NotFound = 0
While ((Repeat1__numRows <> 0) AND (NOT rsGetItems.EOF)) 


if (rsGetItems.Fields.Item("ProductDetailID").Value) = Clng(Request.Form("Item")) then 
NotFound = 1

else 

end if 


  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetItems.MoveNext()

Wend
If NotFound = 0 then
response.write "<font size=3 color=red><b>Item #"+ request.form("Item") +" not in bin</b></font><br/>"
End if

rsGetItems.Requery()




' do only if it finds item match
While ((Repeat1__numRows <> 0) AND (NOT rsGetItems.EOF))

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

  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetItems.MoveNext()
Wend
rsGetItems.Requery()

end if ' if the form field item is empty %>

           <div class="ShowItems">       
  <% 
DetectComplete = 0
While ((Repeat1__numRows <> 0) AND (NOT rsGetItems.EOF)) 
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
%>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetItems.MoveNext()
Wend 
%>
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
