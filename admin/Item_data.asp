<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetDatesSold__MMColParam
rsGetDatesSold__MMColParam = "1"
If (Request.QueryString("DetailID") <> "") Then 
  rsGetDatesSold__MMColParam = Request.QueryString("DetailID")
End If
%>
<%
Dim rsGetDatesSold
Dim rsGetDatesSold_numRows

Set rsGetDatesSold = Server.CreateObject("ADODB.Recordset")
rsGetDatesSold.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetDatesSold.Source = "SELECT * FROM dbo.QRY_DatesItemSold WHERE DetailID = " + Replace(rsGetDatesSold__MMColParam, "'", "''") + " ORDER BY ID DESC"
rsGetDatesSold.CursorType = 0
rsGetDatesSold.CursorLocation = 2
rsGetDatesSold.LockType = 1
rsGetDatesSold.Open()

rsGetDatesSold_numRows = 0
%>
<%
Dim rsGetItem__MMColParam
rsGetItem__MMColParam = "1"
If (Request.QueryString("DetailID") <> "") Then 
  rsGetItem__MMColParam = Request.QueryString("DetailID")
End If
%>
<%
Dim rsGetItem
Dim rsGetItem_numRows

Set rsGetItem = Server.CreateObject("ADODB.Recordset")
rsGetItem.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetItem.Source = "SELECT title, ProductDetail1, ProductID, ProductDetailID FROM dbo.inventory WHERE ProductDetailID = " + Replace(rsGetItem__MMColParam, "'", "''") + ""
rsGetItem.CursorType = 0
rsGetItem.CursorLocation = 2
rsGetItem.LockType = 1
rsGetItem.Open()

rsGetItem_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetDatesSold_numRows = rsGetDatesSold_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsGetDatesSold_total
Dim rsGetDatesSold_first
Dim rsGetDatesSold_last

' set the record count
rsGetDatesSold_total = rsGetDatesSold.RecordCount

' set the number of rows displayed on this page
If (rsGetDatesSold_numRows < 0) Then
  rsGetDatesSold_numRows = rsGetDatesSold_total
Elseif (rsGetDatesSold_numRows = 0) Then
  rsGetDatesSold_numRows = 1
End If

' set the first and last displayed record
rsGetDatesSold_first = 1
rsGetDatesSold_last  = rsGetDatesSold_first + rsGetDatesSold_numRows - 1

' if we have the correct record count, check the other stats
If (rsGetDatesSold_total <> -1) Then
  If (rsGetDatesSold_first > rsGetDatesSold_total) Then
    rsGetDatesSold_first = rsGetDatesSold_total
  End If
  If (rsGetDatesSold_last > rsGetDatesSold_total) Then
    rsGetDatesSold_last = rsGetDatesSold_total
  End If
  If (rsGetDatesSold_numRows > rsGetDatesSold_total) Then
    rsGetDatesSold_numRows = rsGetDatesSold_total
  End If
End If
%>

<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsGetDatesSold_total = -1) Then

  ' count the total records by iterating through the recordset
  rsGetDatesSold_total=0
  While (Not rsGetDatesSold.EOF)
    rsGetDatesSold_total = rsGetDatesSold_total + 1
    rsGetDatesSold.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsGetDatesSold.CursorType > 0) Then
    rsGetDatesSold.MoveFirst
  Else
    rsGetDatesSold.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsGetDatesSold_numRows < 0 Or rsGetDatesSold_numRows > rsGetDatesSold_total) Then
    rsGetDatesSold_numRows = rsGetDatesSold_total
  End If

  ' set the first and last displayed record
  rsGetDatesSold_first = 1
  rsGetDatesSold_last = rsGetDatesSold_first + rsGetDatesSold_numRows - 1
  
  If (rsGetDatesSold_first > rsGetDatesSold_total) Then
    rsGetDatesSold_first = rsGetDatesSold_total
  End If
  If (rsGetDatesSold_last > rsGetDatesSold_total) Then
    rsGetDatesSold_last = rsGetDatesSold_total
  End If

End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Detailed item data</title>
<link href="../includes/nav.css" rel="stylesheet" type="text/css" />
</head>

<body class="mainbkgd">
<!--#include file="admin_header.asp"-->
<span class="adminheader">Detailed data for <%=(rsGetItem.Fields.Item("title").Value)%>
  &nbsp;<%=(rsGetItem.Fields.Item("ProductDetail1").Value)%></span><br><br>
<table width="25%" border="0" cellspacing="1" cellpadding="2">
  <tr class="LeftNavHeaders">
    <td colspan="2" bgcolor="#000000">Items  sold: <%=(rsGetDatesSold_total)%></td>
  </tr>
  <tr class="LeftNavHeaders">
    <td width="50%" bgcolor="#000000">Date sold </td>
    <td width="50%" bgcolor="#000000">Invoice</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetDatesSold.EOF)) 
%>
    <tr class="pricegauge">
      <td width="50%" bgcolor="#EAEAEA" class="materialText"><%=FormatDateTime((rsGetDatesSold.Fields.Item("date_order_placed").Value), 2)%></td>
      <td width="50%" bgcolor="#EAEAEA"><a href="invoice.asp?ID=<%=(rsGetDatesSold.Fields.Item("ID").Value)%>" target="_blank"><%=(rsGetDatesSold.Fields.Item("ID").Value)%></a></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetDatesSold.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
rsGetDatesSold.Close()
Set rsGetDatesSold = Nothing
%>
<%
rsGetItem.Close()
Set rsGetItem = Nothing
%>
