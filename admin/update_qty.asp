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
rsUpdate.Source = "SELECT qty, DetailID, InvoiceID, title, ProductDetail1  FROM QRY_OrderDetails  WHERE ID = " + Request.QueryString("ID") + ""
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
<span class="adminheader">DEDUCT quantities for order #<%=(rsUpdate.Fields.Item("InvoiceID").Value)%></span>
<form action="update_qty.asp" method="post" name="updateQTY" id="updateQTY">
  <p> <span class="faqs">
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsUpdate.EOF)) 
%>
    </MM:DECORATION></MM_REPEATEDREGION>
  </span>
    <MM_REPEATEDREGION NAME="Repeat1" SOURCE="rsUpdate"><MM:DECORATION OUTLINE="Repeat" OUTLINEID=1><span class="pricegauge"><%=(rsUpdate.Fields.Item("qty").Value)%>
      <input name="update_qty" type="hidden" value="<%=(rsUpdate.Fields.Item("qty").Value)%>">
      | <%=(rsUpdate.Fields.Item("title").Value)%>&nbsp;<%=(rsUpdate.Fields.Item("ProductDetail1").Value)%>
      <input name="inventory" type="hidden" value="<%=(rsUpdate.Fields.Item("detailID").Value)%>">
    </span></MM:DECORATION></MM_REPEATEDREGION>
    <span class="faqs">
        <MM_REPEATEDREGION NAME="Repeat1" SOURCE="rsUpdate"><MM:DECORATION OUTLINE="Repeat" OUTLINEID=1><br>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsUpdate.MoveNext()
Wend
%>
    <br>
    </span><input type="submit" name="Submit" value="Submit">
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

If UBound(inventoryItems) <> UBound(quantities) Then 
    Response.Write "BUG!  Array sizes don't match!" 
    Response.End 
End If 

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
For i = 0 To UBound(inventoryItems) 
commUpdate.CommandText = "UPDATE ProductDetails SET qty = qty - " & quantities(i) _ 
           & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = " & inventoryItems(i) 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()
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

