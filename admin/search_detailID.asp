<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsSearch__MMColParam
rsSearch__MMColParam = "1"
If (Request.Form("DetailID") <> "") Then 
  rsSearch__MMColParam = Request.Form("DetailID")
End If
%>
<%
Dim rsSearch
Dim rsSearch_numRows

Set rsSearch = Server.CreateObject("ADODB.Recordset")
rsSearch.ActiveConnection = MM_bodyartforms_sql_STRING
rsSearch.Source = "SELECT title, ProductDetail1, ProductID, ProductDetailID, DetailPrice FROM dbo.inventory WHERE ProductDetailID = " + Replace(rsSearch__MMColParam, "'", "''") + ""
rsSearch.CursorLocation = 3 'adUseClient
rsSearch.LockType = 1 'Read-only records
rsSearch.Open()

rsSearch_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsSearch_numRows = rsSearch_numRows + Repeat1__numRows
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Detail ID search</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#666699" text="#CCCCCC" link="#FFCC00" vlink="#FFCC00" alink="#FFFF00" topmargin="0">

  <!--#include file="admin_header.asp"-->
<span class="adminheader">Detail ID easy search </span> <br>
<br>
<table width="60%" border="0" cellspacing="1" cellpadding="3">
  <tr>
    <td bgcolor="#000000">&nbsp;</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsSearch.EOF)) 
%>
    <tr>
      <td bgcolor="#41426D"><span class="pricegauge">&nbsp;(<%=(rsSearch.Fields.Item("ProductDetailID").Value)%>) &nbsp;<%=(rsSearch.Fields.Item("title").Value)%> - <%=(rsSearch.Fields.Item("ProductDetail1").Value)%>  &nbsp;&nbsp;<a href="addtoorder.asp?ProductID=<%=(rsSearch.Fields.Item("ProductID").Value)%>&DetailID=<%=(rsSearch.Fields.Item("ProductDetailID").Value)%>&Price=<%=(rsSearch.Fields.Item("DetailPrice").Value)%>" class="productnav">Add</a></span></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsSearch.MoveNext()
Wend
%>

</table>
<p><br>

<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
<%
rsSearch.Close()
Set rsSearch = Nothing
%>
