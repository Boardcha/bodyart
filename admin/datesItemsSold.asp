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
rsGetDatesSold.CursorLocation = 3 'adUseClient
rsGetDatesSold.LockType = 1 'Read-only records
rsGetDatesSold.Open()

rsGetDatesSold_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetDatesSold_numRows = rsGetDatesSold_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="../includes/nav.css" rel="stylesheet" type="text/css" />
</head>

<body class="mainbkgd">
<table width="100%" border="0" cellspacing="1" cellpadding="2">
  <tr>
    <td width="30%" bgcolor="#000000" class="adminheader">Date sold </td>
    <td width="50%" bgcolor="#000000" class="adminheader">Invoice</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetDatesSold.EOF)) 
%>
    <tr class="pricegauge">
      <td width="30%" bgcolor="#EAEAEA" class="materialText"><%=FormatDateTime((rsGetDatesSold.Fields.Item("date_order_placed").Value), 2)%></td>
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
