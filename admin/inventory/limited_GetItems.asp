<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetItems__MMColParam
rsGetItems__MMColParam = "1"
If (Request.Form("Bin") <> "") Then 
  rsGetItems__MMColParam = Request.Form("Bin")
  
set SetScanTime = Server.CreateObject("ADODB.Command")
SetScanTime.ActiveConnection = MM_bodyartforms_sql_STRING
SetScanTime.CommandText = "UPDATE TBL_BinNumbers SET BinCountDate = '" & now() & "' WHERE BinNumberID = " & Request.Form("Bin") 
SetScanTime.Execute()
  
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
<link href="../../includes/nav.css" rel="stylesheet" type="text/css" />
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
  <input type="text" name="Item" id="Item" />
  <input name="Bin" type="hidden" id="Bin" value="<%= Request.Form("Bin") %>" /> 
  <strong>BIN # <%= Request.Form("Bin") %>  </strong>
</form>
<br/>
           <% If Not rsGetItems.EOF Or Not rsGetItems.BOF Then %>

<%
' check to see if item scanned matches a product code in order
 if request.form("Item") <> "" then

LoopOnlyOnce = 1
While ((Repeat1__numRows <> 0) AND (NOT rsGetItems.EOF)) 
%>
<% 
if (rsGetItems.Fields.Item("ProductDetailID").Value) = CLng(Request.Form("Item")) then 


LoopOnlyOnce = 2

set SetItemQty = Server.CreateObject("ADODB.Command")
SetItemQty.ActiveConnection = MM_bodyartforms_sql_STRING
SetItemQty.CommandText = "UPDATE ProductDetails SET Inventory_TimesScanned = Inventory_TimesScanned + 1, Date_InventoryCount = '" & now() & "' WHERE ProductDetailID = " & rsGetItems.Fields.Item("ProductDetailID").Value 
SetItemQty.Execute()

end if
%>

  <%  
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
        <% if rsGetItems.Fields.Item("Inventory_TimesScanned").Value <> rsGetItems.Fields.Item("qty").Value then ' only show if there are items left to be scanned 
		
DetectComplete = 1 %>
  <% if (rsGetItems.Fields.Item("qty").Value) - (rsGetItems.Fields.Item("Inventory_TimesScanned").Value) < 0 then %>
  <strong><font color="#0033FF">OVER <%=((rsGetItems.Fields.Item("qty").Value) - (rsGetItems.Fields.Item("Inventory_TimesScanned").Value)) * -1%></font></strong>
  <% else %><%=(rsGetItems.Fields.Item("qty").Value) - (rsGetItems.Fields.Item("Inventory_TimesScanned").Value)%><% end if %>
&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;<strong><%=(rsGetItems.Fields.Item("ProductDetailID").Value)%></strong>&nbsp;&nbsp;&nbsp;<%=(rsGetItems.Fields.Item("title").Value)%>&nbsp;<%=(rsGetItems.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetItems.Fields.Item("Length").Value)%> <%=(rsGetItems.Fields.Item("ProductDetail1").Value)%>&nbsp;&nbsp;<% if (rsGetItems.Fields.Item("DateLastPurchased").Value) > now()-1 then %>
 <strong><font color="#CC0000">SOLD RECENTLY</font></strong>
<% end if %><br /><% end if %> 
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