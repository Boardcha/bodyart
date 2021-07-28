<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/BAFDatesSold.asp" -->
<%
Dim rsUpdate__MMColParam
rsUpdate__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsUpdate__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsUpdate
Dim rsUpdate_numRows

Set rsUpdate = Server.CreateObject("ADODB.Recordset")
rsUpdate.ActiveConnection = MM_BAFDatesSold_STRING
rsUpdate.Source = "SELECT * FROM TBLDates WHERE DetailID = " + Replace(rsUpdate__MMColParam, "'", "''") + " ORDER BY DateSold DESC"
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
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Update inventory quantities</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#666699" text="#CCCCCC" link="#FFCC00" vlink="#FFCC00">
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsUpdate.EOF)) 
%>
<font size="2" face="Verdana"><span class="smallestfont"><%=(rsUpdate.Fields.Item("DateSold").Value)%></span><br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsUpdate.MoveNext()
Wend
%>

<p>&nbsp;</p>
</body>
</html>
<%
rsUpdate.Close()
Set rsUpdate = Nothing
%>
