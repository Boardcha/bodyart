<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Update inventory quantities</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#666699" text="#CCCCCC" link="#FFCC00" vlink="#FFCC00">
<% if Request.form("process") <> "yes" then %>
<%
Dim rsUpdate
Dim rsUpdate_numRows

Set rsUpdate = Server.CreateObject("ADODB.Recordset")
rsUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
rsUpdate.Source = "SELECT qty, DetailID, InvoiceID, jewelry, ActiveDetail, ActiveMain FROM QRY_OrderDetails WHERE ID = " + Request.QueryString("ID") + ""
rsUpdate.CursorLocation = 3 'adUseClient
rsUpdate.LockType = 1 'Read-only records
rsUpdate.Open()

rsUpdate_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsUpdate_numRows = rsUpdate_numRows + Repeat1__numRows
%>
<span class="adminheader">PUT ITEMS BACK IN STOCK for  #<%=(rsUpdate.Fields.Item("InvoiceID").Value)%></span>
<form action="update_qty_ADDBACK.asp" method="post" name="updateQTY" id="updateQTY">
  <p> <span class="faqs">
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsUpdate.EOF)) 
%>
    <%=(rsUpdate.Fields.Item("qty").Value)%>
    <input name="update_qty" type="hidden" value="<%=(rsUpdate.Fields.Item("qty").Value)%>">
    | <%=(rsUpdate.Fields.Item("detailID").Value)%>
      <input name="inventory" type="hidden" value="<%=(rsUpdate.Fields.Item("detailID").Value)%>">
      <input name="jewelry" type="hidden" value="<%=(rsUpdate.Fields.Item("jewelry").Value)%>">
      <input name="ActiveMain" type="hidden" value="<%=(rsUpdate.Fields.Item("ActiveMain").Value)%>">
      <input name="ActiveDetail" type="hidden" value="<%=(rsUpdate.Fields.Item("ActiveDetail").Value)%>">
      <br>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsUpdate.MoveNext()
Wend
%>
<br>
  </span> </p>
  <p>
    <input type="submit" name="Submit" value="Submit">
    <input name="process" type="hidden" value="yes">
  </p>
</form>
<p><font size="2" face="Verdana">
<%
rsUpdate.Close()
Set rsUpdate = Nothing
%>
<% end if %>
<% if Request.form("process") = "yes" then %>
<% 
temp = Replace( Request.Form("inventory"), "'", "''" ) 
inventoryItems = Split( temp, ", " ) 
temp = Replace( Request.Form("update_qty"), "'", "''" ) 
quantities = Split( temp, ", " ) 
temp = Replace( Request.Form("jewelry"), "'", "''" ) 
jewelry = Split( temp, ", " ) 
temp = Replace( Request.Form("ActiveMain"), "'", "''" ) 
ActiveMain = Split( temp, ", " )
temp = Replace( Request.Form("ActiveDetail"), "'", "''" ) 
ActiveDetail = Split( temp, ", " )  

If UBound(inventoryItems) <> UBound(quantities) Then 
    Response.Write "BUG!  Array sizes don't match!" 
    Response.End 
End If 

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
For i = 0 To UBound(inventoryItems) 

commUpdate.CommandText = "UPDATE ProductDetails SET qty = qty + " & quantities(i) _ 
           & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = " & inventoryItems(i) 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()

if jewelry(i) <> "save" then
	
	if ActiveDetail(i) = 0 then
	
		set UpdateDetailActive = Server.CreateObject("ADODB.Command")
		UpdateDetailActive.ActiveConnection = MM_bodyartforms_sql_STRING
		UpdateDetailActive.CommandText = "UPDATE ProductDetails SET active = 1 WHERE ProductDetailID = " & inventoryItems(i) 
   		 ' comment out next line AFTER IT WORKS 
   		 'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
		UpdateDetailActive.Execute()
	
	end if ' make detail active if it was inactive
	
	if ActiveMain(i) = 0 then
	
		set UpdateMainActive = Server.CreateObject("ADODB.Command")
		UpdateMainActive.ActiveConnection = MM_bodyartforms_sql_STRING
		UpdateMainActive.CommandText = "UPDATE QRY_OrderDetails SET ActiveMain = 1 WHERE ProductDetailID = " & inventoryItems(i) 
   		 ' comment out next line AFTER IT WORKS 
   		 'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
		UpdateMainActive.Execute()
	
	end if ' make main product active if it was inactive
	
end if ' jewelry isn't in the saved category

Next
%>
<% end if %>
</font></p>
<p>
  <% if Request.form("process") = "yes" then %>
Updates have been made
<script language="JavaScript" type="text/JavaScript">
  <!--
    self.close();
  // -->
</script>
    <% end if %>
</p>
</body>
</html>