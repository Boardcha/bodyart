<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT ID FROM sent_items WHERE  (shipped = N'Pending...') AND (ship_code IS NULL) AND (date_order_placed < CONVERT(DATETIME, GETDATE() - 45, 102)) ORDER BY ID ASC"
'rsGetRecords.Source = "SELECT ID FROM sent_items WHERE (shipped = N'Pending...') AND (ship_code <> 'paid') AND  (date_order_placed <= '" & date()-45& "') ORDER BY ID ASC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
rsGetRecords_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetRecords_numRows = rsGetRecords_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsGetRecords_total
Dim rsGetRecords_first
Dim rsGetRecords_last

' set the record count
rsGetRecords_total = rsGetRecords.RecordCount

' set the number of rows displayed on this page
If (rsGetRecords_numRows < 0) Then
  rsGetRecords_numRows = rsGetRecords_total
Elseif (rsGetRecords_numRows = 0) Then
  rsGetRecords_numRows = 1
End If

' set the first and last displayed record
rsGetRecords_first = 1
rsGetRecords_last  = rsGetRecords_first + rsGetRecords_numRows - 1

' if we have the correct record count, check the other stats
If (rsGetRecords_total <> -1) Then
  If (rsGetRecords_first > rsGetRecords_total) Then
    rsGetRecords_first = rsGetRecords_total
  End If
  If (rsGetRecords_last > rsGetRecords_total) Then
    rsGetRecords_last = rsGetRecords_total
  End If
  If (rsGetRecords_numRows > rsGetRecords_total) Then
    rsGetRecords_numRows = rsGetRecords_total
  End If
End If
%>

<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsGetRecords_total = -1) Then

  ' count the total records by iterating through the recordset
  rsGetRecords_total=0
  While (Not rsGetRecords.EOF)
    rsGetRecords_total = rsGetRecords_total + 1
    rsGetRecords.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsGetRecords.CursorType > 0) Then
    rsGetRecords.MoveFirst
  Else
    rsGetRecords.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsGetRecords_numRows < 0 Or rsGetRecords_numRows > rsGetRecords_total) Then
    rsGetRecords_numRows = rsGetRecords_total
  End If

  ' set the first and last displayed record
  rsGetRecords_first = 1
  rsGetRecords_last = rsGetRecords_first + rsGetRecords_numRows - 1
  
  If (rsGetRecords_first > rsGetRecords_total) Then
    rsGetRecords_first = rsGetRecords_total
  End If
  If (rsGetRecords_last > rsGetRecords_total) Then
    rsGetRecords_last = rsGetRecords_total
  End If

End If
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Clean orders out of DB</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body topmargin="0" class="mainbkgd">
<!--#include file="admin_header.asp"-->
<p><span class="adminheader"><%=(rsGetRecords_total)%> orders to delete </span><br>
This will delete ALL unpaid orders over 45 days old. Press submit to continue...</p>
<p>After submitting this page will look exactly the same except the status bar will be &quot;working/thinking&quot;. After a while the page may timeout. If this happens just click back to the page and submit again to do more. </p>
<form action="cleanDB.asp?delete=yes" method="post" name="FRMDelete" id="FRMDelete">
  <input type="submit" name="Submit" value="Submit">
</form>
<% if request.querystring("delete") = "yes" then %>
<% 
Do until rsGetRecords.EOF

set DeleteItem = Server.CreateObject("ADODB.Command")
DeleteItem.ActiveConnection = MM_bodyartforms_sql_STRING
DeleteItem.CommandText = "DELETE FROM TBL_OrderSummary WHERE InvoiceID = " & (rsGetRecords.Fields.Item("ID").Value) & ""
DeleteItem.CommandType = 1
DeleteItem.CommandTimeout = 0
DeleteItem.Prepared = true
DeleteItem.Execute()


set DeleteInvoice = Server.CreateObject("ADODB.Command")
DeleteInvoice.ActiveConnection = MM_bodyartforms_sql_STRING
DeleteInvoice.CommandText = "DELETE FROM sent_items WHERE ID = " & (rsGetRecords.Fields.Item("ID").Value) & ""
DeleteInvoice.CommandType = 1
DeleteInvoice.CommandTimeout = 0
DeleteInvoice.Prepared = true
DeleteInvoice.Execute()

rsGetRecords.MoveNext
Loop

Response.Redirect("cleanDB.asp")

%>
<% end if %>
<p>&nbsp; </p>
</body>
</html>
<% rsGetRecords.Close()
set rsGetRecords = Nothing
%>

