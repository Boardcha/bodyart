<%@LANGUAGE="VBSCRIPT"%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

Dim rsGetRecords__MMColParam2
rsGetRecords__MMColParam2 = "date_order_placed"
if (Request("sortby") <> "") then rsGetRecords__MMColParam2 = Request("sortby")
%>
<%
set DataConn = Server.CreateObject("ADODB.connection")
DataConn.Open MM_bodyartforms_sql_STRING ' CONNECTION STRING FOR ALL PROCEDURES


set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = DataConn
rsGetRecords.Source = "SELECT *  FROM QRY_Backorders2 WHERE shipped <> N'SHIPPING BACKORDER' AND customorder <> N'yes' ORDER BY " + Replace(rsGetRecords__MMColParam2, "'", "''") + " ASC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
rsGetRecords_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
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

<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<html>
<head>
<title>Backorders</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="px-2">
<h5 class="mt-3"><%=(rsGetRecords_total)%> backorders</h5>

<table class="table table-striped table-hover table-sm">

  <thead class="thead-dark">
  <tr> 
    <th width="15%"></th>
    <th width="20%"><a href="backorders.asp?sortby=customer_first">Sort by first name</a></th>
    <th width="50%">Items backordered</th>
    <th width="15%" class="text-right"><a href="backorders.asp?sortby=date_order_placed">Sort by date</a></th>
  </tr>
</thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetRecords.EOF)) 
%>
    <tr <% if rsGetRecords.Fields.Item("QtyInStock").Value > 0 then %>class="table-success"<% end if %>> 
      <td>
        <a class="btn btn-secondary btn-sm mr-2" href="invoice.asp?<%= MM_keepURL & MM_joinChar(MM_keepURL) & "ID=" & rsGetRecords.Fields.Item("ID").Value %>"><%=(rsGetRecords.Fields.Item("ID").Value)%></a>
        <a class="btn btn-secondary btn-sm" href="order history.asp?var_first=<%=(rsGetRecords.Fields.Item("customer_first").Value)%>&var_last=<%=(rsGetRecords.Fields.Item("customer_last").Value)%>">History</a>
      </td>
      <td>
        <%=(rsGetRecords.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsGetRecords.Fields.Item("customer_last").Value)%><% if (rsGetRecords.Fields.Item("customer_ID").Value) <> 0 then%>&nbsp;&nbsp;(Registered)<% end if %></td>
      <td>
          <a class="btn btn-warning font-weight-bold py-0 px-2 mr-1" href="invoice.asp?ID=<%=(rsGetRecords.Fields.Item("ID").Value)%>&bo_item=<%=(rsGetRecords.Fields.Item("OrderDetailID").Value)%>">ON BO</a>
          <% if trim(rsGetRecords.Fields.Item("notes").Value) <> "" then %>
          <span class="bg-transparent font-weight-bold text-dark py-0 px-2 mr-2 border border-dark rounded">
          <%= rsGetRecords.Fields.Item("notes").Value %>
        </span>
          <% end if %>
        <% if rsGetRecords.Fields.Item("QtyInStock").Value > 0 then %>
        <span class="font-weight-bold text-success"><%=(rsGetRecords.Fields.Item("QtyInStock").Value)%> IN STOCK</span>
      <% end if %> <%=(rsGetRecords.Fields.Item("qty").Value)%> | <a href="product-edit.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%>&info=less" class="productnav"><%=(rsGetRecords.Fields.Item("title").Value)%></a>&nbsp; <%=(rsGetRecords.Fields.Item("Gauge").Value)%>&nbsp; <%=(rsGetRecords.Fields.Item("Length").Value)%>&nbsp; <%=(rsGetRecords.Fields.Item("ProductDetail1").Value)%>&nbsp;&nbsp;$<%=(rsGetRecords.Fields.Item("item_price").Value) * (rsGetRecords.Fields.Item("qty").Value)%>

    </td>
      <td class="text-right">
        <%= FormatDateTime(rsGetRecords.Fields.Item("date_order_placed").Value,2)%>
      </td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetRecords.MoveNext()
Wend
%>
</table>


</div><!-- content div-->
</body>
</html>
<%
rsGetRecords.Close()
%>
