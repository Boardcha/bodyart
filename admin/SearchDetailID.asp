<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetProduct__MMColParam
rsGetProduct__MMColParam = "1"
If (Request("DetailID") <> "") Then 
  rsGetProduct__MMColParam = Request("DetailID")
End If

if request.form("sku") <> "" then
	var_sku = request.form("sku")
else
	var_sku = "test100002"
end if
%>
<%
Dim rsGetProduct
Dim rsGetProduct_cmd
Dim rsGetProduct_numRows

Set rsGetProduct_cmd = Server.CreateObject ("ADODB.Command")
rsGetProduct_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetProduct_cmd.CommandText = "SELECT ProductDetailID, ProductID FROM dbo.ProductDetails WHERE ProductDetailID = ? OR detail_code = ?" 
rsGetProduct_cmd.Prepared = true
rsGetProduct_cmd.Parameters.Append rsGetProduct_cmd.CreateParameter("param1", 5, 1, -1, rsGetProduct__MMColParam) ' adDouble
rsGetProduct_cmd.Parameters.Append(rsGetProduct_cmd.CreateParameter("sku",200,1,15,var_sku))

Set rsGetProduct = rsGetProduct_cmd.Execute
rsGetProduct_numRows = 0
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Search detail ID</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body  topmargin="0" class="mainbkgd" >
  <!--#include file="admin_header.asp"-->
&nbsp;
<p class="adminheader"> Searching details... </p>
<% If Not rsGetProduct.EOF Or Not rsGetProduct.BOF Then %>
<p class="adminheader"><% Response.Redirect "product-edit.asp?ProductID="& (rsGetProduct.Fields.Item("ProductID").Value) &"&info=less" %></p>
  <% End If ' end Not rsGetProduct.EOF Or NOT rsGetProduct.BOF %>
<% If rsGetProduct.EOF And rsGetProduct.BOF Then %>
  <p class="adminheader">No product found</p>
  <% End If ' end rsGetProduct.EOF And rsGetProduct.BOF %>
</body>
</html>
<%
rsGetProduct.Close()
Set rsGetProduct = Nothing
%>
