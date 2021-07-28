<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.Buffer = True %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Set conn = Server.CreateObject ("ADODB.Connection")
conn.open = MM_bodyartforms_sql_STRING

SortBy = Request("SortBy")

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = conn
objCmd.CommandText = "SELECT * FROM dbo.QRY_PORestock WHERE PurchaseOrderID = ? ORDER BY " + SortBy + "" 
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Request("ID"))
Set rsGetRestockItems = objCmd.Execute		  
Set objCmd = Nothing

		  rsGetRestockItems_numRows = 0
		  Dim RepeatRestock__numRows
		  Dim RepeatRestock__index
		  
		  RepeatRestock__numRows = -1
		  RepeatRestock__index = 0
		  rsGetRestockItems_numRows = rsGetRestockItems_numRows + RepeatRestock__numRows

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = conn
objCmd.CommandText = "SELECT PurchaseOrderID, Brand FROM dbo.TBL_PurchaseOrders WHERE PurchaseOrderID = ?" 
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Request("ID"))
Set rsGetCurrentPOInfo = objCmd.Execute
Set objCmd = Nothing


If Request.Form("Brand") <> "" Then

		  'Create a new empty purchase order
		  set objCmd = Server.CreateObject("ADODB.Command")
		  objCmd.ActiveConnection = conn
		  objCmd.CommandText = "INSERT INTO TBL_PurchaseOrders (DateOrdered, Brand) VALUES ('"& date() &"', '"& request.form("Brand") & "')" 
		  objCmd.Execute()
		  Set objCmd = Nothing		


		  'Find the last purchase order ID in table
		  Set objCmd = Server.CreateObject ("ADODB.Command")
		  objCmd.ActiveConnection = conn
		  objCmd.CommandText = "SELECT TOP 1 PurchaseOrderID FROM dbo.TBL_PurchaseOrders ORDER BY PurchaseOrderID DESC" 
		  objCmd.Prepared = true	  
		  Set rsGetLastPO = objCmd.Execute
		  Set objCmd = Nothing



		  temp = Replace( Request.Form("Checkbox"), "'", "''" ) 
		  varID = Split( temp, ", " ) 
		  
		  'Response.write temp
		  
		  For i = 0 To UBound(varID)
		  	SqlText = "ProductDetailID = " & varID(i) & " OR "
			SqlText2 = SqlText2 + SqlText
		  Next
		  
		  'Response.write SqlText2
		  
		  For i = 0 To UBound(varID)
		  
				set comm = Server.CreateObject("ADODB.Command")
				comm.ActiveConnection = MM_bodyartforms_sql_STRING
				comm.CommandText = "UPDATE dbo.ProductDetails SET PurchaseOrderID="& rsGetLastPO.Fields.Item("PurchaseOrderID").Value &" WHERE " & SqlText2 & " ProductDetailID = 0" 
				comm.Execute()
		  
		 Next
				Set comm = Nothing
		  
		  Response.Redirect "PurchaseOrders_PutInStock.asp?ID=" & rsGetLastPO.Fields.Item("PurchaseOrderID").Value & "&SortBy=ProductDetailID ASC"
		  
		  rsGetLastPO.Close()
		  Set rsGetLastPO = Nothing

End if ' only update if form has been submitted
%>
<html>
<head>

<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Split orders</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body topmargin="0" class="MainBkgdColor">
<!--#include file="admin_header.asp"-->
<span class="adminheader">Split order out into two separate orders<br>
</span>&nbsp;&nbsp;
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td colspan="2"> <div align="left" class="faqs">
      <% If Not rsGetRestockItems.EOF Or Not rsGetRestockItems.BOF Then %>
<form METHOD="POST" ACTION="PurchaseOrders_Split.asp" name="FRM_UpdateStock" id="FRM_UpdateStock">          
  <table width="60%" border="0" cellspacing="1" cellpadding="6">
            <tr valign="middle" bgcolor="#000000">
              <td bgcolor="#000000" class="faqs"><span class="checkoutHeader">Select the items that you want to split into a new order</span><br>
              <span class="pricegauge">The items that are not selected will remain in the original order you created</span></td>
            </tr>
                  <tr valign="middle" bgcolor="#ececec">
                    <td bgcolor="#ececec"><a href="PurchaseOrders_Split.asp?ID=<%= Request.Querystring("ID") %>&SortBy=ProductDetailID ASC" class="Link_ItemDetails">Sort by #</a>&nbsp; |&nbsp;&nbsp;
                      <a href="PurchaseOrders_Split.asp?ID=<%= Request.Querystring("ID") %>&SortBy=title ASC" class="Link_ItemDetails">Sort by name</a>
</td>
                  </tr>              <% 
While ((RepeatRestock__numRows <> 0) AND (NOT rsGetRestockItems.EOF)) 
%>

                <tr valign="middle" bgcolor="#ececec">
                  <td bgcolor="#ececec"><p><font color="#999999">
                  <% upd_rsGetRestockItems = upd_rsGetRestockItems+1 %>
                  <input name="Checkbox" type="checkbox" value="<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%>">
                  &nbsp;<a href="product-edit.asp?ProductID=<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>&info=less" target="_blank" class="EditSelect_Links"><strong><%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%></strong></a>&nbsp;&nbsp; </font><span class="materialText"><a href="product-edit.asp?ProductID=<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>&info=less" target="_blank" class="EditSelect_Links"><%=(rsGetRestockItems.Fields.Item("title").Value)%></a> <%=(rsGetRestockItems.Fields.Item("Gauge").Value)%> <%=(rsGetRestockItems.Fields.Item("Length").Value)%> <%=(rsGetRestockItems.Fields.Item("ProductDetail1").Value)%></span></p></td>
              </tr>
                <% 
  RepeatRestock__index=RepeatRestock__index+1
  RepeatRestock__numRows=RepeatRestock__numRows-1
  rsGetRestockItems.MoveNext()
Wend
%>
          </table>
  <p>
    <input type="submit" name="button" id="button" value="Split to new order">
    <input name="Brand" type="hidden" id="Brand" value="<%=(rsGetCurrentPOInfo.Fields.Item("Brand").Value)%>">
    <input name="SortBy" type="hidden" id="SortBy" value="<%= Request.Querystring("SortBy") %>">
    <input name="ID" type="hidden" id="ID" value="<%= Request.Querystring("ID") %>">
  </p>
</form>
          <% End If ' end Not rsGetRestockItems.EOF Or NOT rsGetRestockItems.BOF %>
      <% If rsGetRestockItems.EOF And rsGetRestockItems.BOF Then %>
        <p class="adminheader"><font color="#FFFF00">No order to display</font></p>
        <% End If ' end rsGetRestockItems.EOF And rsGetRestockItems.BOF %>
</div></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
rsGetRestockItems.Close()
Set rsGetRestockItems = Nothing

rsGetCurrentPOInfo.Close()
Set rsGetCurrentPOInfo = Nothing
%>
