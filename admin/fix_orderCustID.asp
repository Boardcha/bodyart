<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetOrders
Dim rsGetOrders_cmd
Dim rsGetOrders_numRows

Set rsGetOrders_cmd = Server.CreateObject ("ADODB.Command")
rsGetOrders_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrders_cmd.CommandText = "SELECT customer_ID, ID, customer_first, customer_last, email FROM dbo.sent_items WHERE customer_ID = 19477 ORDER BY ID ASC" 
rsGetOrders_cmd.Prepared = true
rsGetOrders_cmd.Parameters.Append rsGetOrders_cmd.CreateParameter("param1", 5, 1, -1, rsGetOrders__MMColParam) ' adDouble

Set rsGetOrders = rsGetOrders_cmd.Execute
rsGetOrders_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 1000
Repeat1__index = 0
rsGetOrders_numRows = rsGetOrders_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsGetOrders_total
Dim rsGetOrders_first
Dim rsGetOrders_last

' set the record count
rsGetOrders_total = rsGetOrders.RecordCount

' set the number of rows displayed on this page
If (rsGetOrders_numRows < 0) Then
  rsGetOrders_numRows = rsGetOrders_total
Elseif (rsGetOrders_numRows = 0) Then
  rsGetOrders_numRows = 1
End If

' set the first and last displayed record
rsGetOrders_first = 1
rsGetOrders_last  = rsGetOrders_first + rsGetOrders_numRows - 1

' if we have the correct record count, check the other stats
If (rsGetOrders_total <> -1) Then
  If (rsGetOrders_first > rsGetOrders_total) Then
    rsGetOrders_first = rsGetOrders_total
  End If
  If (rsGetOrders_last > rsGetOrders_total) Then
    rsGetOrders_last = rsGetOrders_total
  End If
  If (rsGetOrders_numRows > rsGetOrders_total) Then
    rsGetOrders_numRows = rsGetOrders_total
  End If
End If
%>

<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsGetOrders_total = -1) Then

  ' count the total records by iterating through the recordset
  rsGetOrders_total=0
  While (Not rsGetOrders.EOF)
    rsGetOrders_total = rsGetOrders_total + 1
    rsGetOrders.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsGetOrders.CursorType > 0) Then
    rsGetOrders.MoveFirst
  Else
    rsGetOrders.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsGetOrders_numRows < 0 Or rsGetOrders_numRows > rsGetOrders_total) Then
    rsGetOrders_numRows = rsGetOrders_total
  End If

  ' set the first and last displayed record
  rsGetOrders_first = 1
  rsGetOrders_last = rsGetOrders_first + rsGetOrders_numRows - 1
  
  If (rsGetOrders_first > rsGetOrders_total) Then
    rsGetOrders_first = rsGetOrders_total
  End If
  If (rsGetOrders_last > rsGetOrders_total) Then
    rsGetOrders_last = rsGetOrders_total
  End If

End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
<p><%=(rsGetOrders_total)%></p>
<p>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetOrders.EOF)) 
%>
    <%=(rsGetOrders.Fields.Item("ID").Value)%>&nbsp;<%=(rsGetOrders.Fields.Item("customer_first").Value)%>&nbsp;<%=(rsGetOrders.Fields.Item("customer_last").Value)%>... Upated to customer ID 
   <%
   
Dim rsGetCustomer
Dim rsGetCustomer_cmd
Dim rsGetCustomer_numRows

Set rsGetCustomer_cmd = Server.CreateObject ("ADODB.Command")
rsGetCustomer_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCustomer_cmd.CommandText = "SELECT customer_ID, customer_first, customer_last, email FROM dbo.customers WHERE customer_first = '" + replace(rsGetOrders.Fields.Item("customer_first").Value, "'", "''") + "' AND customer_last = '" + replace(rsGetOrders.Fields.Item("customer_last").Value, "'", "''") + "'" 
rsGetCustomer_cmd.Prepared = true

Set rsGetCustomer = rsGetCustomer_cmd.Execute
rsGetCustomer_numRows = 0
%>
    <% If Not rsGetCustomer.EOF Or Not rsGetCustomer.BOF Then %>
      <%=(rsGetCustomer.Fields.Item("customer_ID").Value)%>

<%
set commUpdate1 = Server.CreateObject("ADODB.Command")
commUpdate1.ActiveConnection = MM_bodyartforms_sql_STRING
CommUpdate1.CommandText = "UPDATE sent_items SET customer_ID = " & rsGetCustomer.Fields.Item("customer_ID").Value  & " WHERE ID = " & rsGetOrders.Fields.Item("ID").Value
commUpdate1.Execute()
%>
      <% End If ' end Not rsGetCustomer.EOF Or NOT rsGetCustomer.BOF %>
    <% If rsGetCustomer.EOF And rsGetCustomer.BOF Then %>
      0
<%
set commUpdate2 = Server.CreateObject("ADODB.Command")
commUpdate2.ActiveConnection = MM_bodyartforms_sql_STRING
CommUpdate2.CommandText = "UPDATE sent_items SET customer_ID = 0 WHERE ID = " & rsGetOrders.Fields.Item("ID").Value
commUpdate2.Execute()
%>
  <% End If ' end rsGetCustomer.EOF And rsGetCustomer.BOF %>
<br />
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetOrders.MoveNext()

rsGetCustomer.Close()
Set rsGetCustomer = Nothing

Wend
%>
</p>
</body>
</html>
<%
rsGetOrders.Close()
Set rsGetOrders = Nothing
%>
