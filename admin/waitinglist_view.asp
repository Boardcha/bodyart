<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


Dim rsShowWaitingList__MMColParam
rsShowWaitingList__MMColParam = "1"
If (Request.QueryString("DetailID") <> "") Then 
  rsShowWaitingList__MMColParam = Request.QueryString("DetailID")
End If
%>
<%
Dim rsShowWaitingList
Dim rsShowWaitingList_numRows

Set rsShowWaitingList = Server.CreateObject("ADODB.Recordset")
rsShowWaitingList.ActiveConnection = MM_bodyartforms_sql_STRING
rsShowWaitingList.Source = "SELECT DetailID, name, email, title, ProductDetail1, (SELECT SUM(waiting_qty) FROM dbo.QRYWaitingList WHERE DetailID = " + Replace(rsShowWaitingList__MMColParam, "'", "''") + ") as total_waiting_qty FROM dbo.QRYWaitingList WHERE DetailID = " + Replace(rsShowWaitingList__MMColParam, "'", "''") + " ORDER BY name ASC"
rsShowWaitingList.CursorLocation = 3 'adUseClient
rsShowWaitingList.LockType = 1 'Read-only records
rsShowWaitingList.Open()

rsShowWaitingList_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsShowWaitingList_numRows = rsShowWaitingList_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsShowWaitingList_total
Dim rsShowWaitingList_first
Dim rsShowWaitingList_last

' set the record count
rsShowWaitingList_total = rsShowWaitingList("total_waiting_qty")

' set the number of rows displayed on this page
If (rsShowWaitingList_numRows < 0) Then
  rsShowWaitingList_numRows = rsShowWaitingList_total
Elseif (rsShowWaitingList_numRows = 0) Then
  rsShowWaitingList_numRows = 1
End If

' set the first and last displayed record
rsShowWaitingList_first = 1
rsShowWaitingList_last  = rsShowWaitingList_first + rsShowWaitingList_numRows - 1

' if we have the correct record count, check the other stats
If (rsShowWaitingList_total <> -1) Then
  If (rsShowWaitingList_first > rsShowWaitingList_total) Then
    rsShowWaitingList_first = rsShowWaitingList_total
  End If
  If (rsShowWaitingList_last > rsShowWaitingList_total) Then
    rsShowWaitingList_last = rsShowWaitingList_total
  End If
  If (rsShowWaitingList_numRows > rsShowWaitingList_total) Then
    rsShowWaitingList_numRows = rsShowWaitingList_total
  End If
End If
%>

<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsShowWaitingList_total = -1) Then

  ' count the total records by iterating through the recordset
  rsShowWaitingList_total=0
  While (Not rsShowWaitingList.EOF)
    rsShowWaitingList_total = rsShowWaitingList_total + 1
    rsShowWaitingList.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsShowWaitingList.CursorType > 0) Then
    rsShowWaitingList.MoveFirst
  Else
    rsShowWaitingList.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsShowWaitingList_numRows < 0 Or rsShowWaitingList_numRows > rsShowWaitingList_total) Then
    rsShowWaitingList_numRows = rsShowWaitingList_total
  End If

  ' set the first and last displayed record
  rsShowWaitingList_first = 1
  rsShowWaitingList_last = rsShowWaitingList_first + rsShowWaitingList_numRows - 1
  
  If (rsShowWaitingList_first > rsShowWaitingList_total) Then
    rsShowWaitingList_first = rsShowWaitingList_total
  End If
  If (rsShowWaitingList_last > rsShowWaitingList_total) Then
    rsShowWaitingList_last = rsShowWaitingList_total
  End If

End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>View waiting list</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
  <!--#include file="admin_header.asp"-->
 <h4><%=(rsShowWaitingList.Fields.Item("title").Value)%>&nbsp;<%=(rsShowWaitingList.Fields.Item("ProductDetail1").Value)%> (<%=(rsShowWaitingList_total)%>)   </h4>

<table class="table table-sm table-borderless">
    <% 
While ((Repeat1__numRows <> 0) AND (NOT rsShowWaitingList.EOF)) 
%>
<tr>
        <td><%=(rsShowWaitingList.Fields.Item("name").Value)%></td>
        <td><%=(rsShowWaitingList.Fields.Item("email").Value)%></td>
    </tr>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsShowWaitingList.MoveNext()
Wend
%>
  </table>
</body>
</html>
<%
rsShowWaitingList.Close()
Set rsShowWaitingList = Nothing
%>
