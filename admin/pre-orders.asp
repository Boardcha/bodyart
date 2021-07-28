<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetOrders__MMColParam
rsGetOrders__MMColParam = "ON ORDER"
If (Request("MM_EmptyValue") <> "") Then 
  rsGetOrders__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsGetOrders
Dim rsGetOrders_numRows

Set rsGetOrders = Server.CreateObject("ADODB.Recordset")
rsGetOrders.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrders.Source = "SELECT ID, shipped, item_description, comments, ship_code FROM dbo.sent_items WHERE shipped = '" + Replace(rsGetOrders__MMColParam, "'", "''") + "' ORDER BY ID ASC"
rsGetOrders.CursorLocation = 3 'adUseClient
rsGetOrders.LockType = 1 'Read-only records
rsGetOrders.Open()

rsGetOrders_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetOrders_numRows = rsGetOrders_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Bodyartforms order</title>
<link href="../includes/styles.css" rel="stylesheet" type="text/css" />
</head>

<body>
<p class="BodyText"><strong>BODYARTFORMS</strong><br />
301 Hester's Crossing STE 206B<br />
  Round Rock, TX  78681<br />
  Cell: 512-417-0003 | service@bodyartforms.com
</p>
<table width="100%" border="1" cellpadding="4" cellspacing="0" bordercolor="#000000">
  <tr>
    <td align="left" valign="top">&nbsp;</td>
    <td align="left" valign="top">&nbsp;</td>
    <td align="left" valign="top">&nbsp;</td>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetOrders.EOF)) 
%>
    <tr>
      <td width="15%" align="left" valign="top"><span class="BodyText"><%=(rsGetOrders.Fields.Item("ID").Value)%></span></td>
      <td width="50%" align="left" valign="top"><span class="BodyText"><%
Dim itm, r, tmp, orig
orig= (rsGetOrders.Fields.Item("item_description").Value)
Set r = New RegExp
itm=orig
r.pattern = "\$[^<]*"
r.Global = True
itm = Replace(itm, ", ", "")
Set r = Nothing
Response.Write itm
%><%=(rsGetOrders.Fields.Item("item_description").Value)%></span></td>
      <td width="35%" align="left" valign="top"><span class="BodyText"><%=(rsGetOrders.Fields.Item("comments").Value)%></span></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetOrders.MoveNext()
Wend
%>
</table>
<p>&nbsp;</p>


</body>
</html>
<%
rsGetOrders.Close()
Set rsGetOrders = Nothing
%>
