<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


Dim rsCheck__MMColParam
rsCheck__MMColParam = "1"
If (Request.QueryString("brand") <> "") Then 
  rsCheck__MMColParam = Request.QueryString("brand")
End If
%>

<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_bodyartforms_sql_STRING
Recordset1.Source = "SELECT * FROM ProductDetails"
Recordset1.CursorLocation = 3 'adUseClient
Recordset1.LockType = 1 'Read-only records
Recordset1.Open()

Recordset1_numRows = 0

set rsCheck = Server.CreateObject("ADODB.Recordset")
rsCheck.ActiveConnection = MM_bodyartforms_sql_STRING
rsCheck.Source = "SELECT qty, ProductDetail1, ProductID, title FROM dbo.inventory  WHERE free <> 0 and item_active = 1 ORDER BY title ASC"
rsCheck.CursorLocation = 3 'adUseClient
rsCheck.LockType = 1 'Read-only records
rsCheck.Open()
rsCheck_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsCheck_numRows = rsCheck_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsCheck_total
Dim rsCheck_first
Dim rsCheck_last

' set the record count
rsCheck_total = rsCheck.RecordCount

' set the number of rows displayed on this page
If (rsCheck_numRows < 0) Then
  rsCheck_numRows = rsCheck_total
Elseif (rsCheck_numRows = 0) Then
  rsCheck_numRows = 1
End If

' set the first and last displayed record
rsCheck_first = 1
rsCheck_last  = rsCheck_first + rsCheck_numRows - 1

' if we have the correct record count, check the other stats
If (rsCheck_total <> -1) Then
  If (rsCheck_first > rsCheck_total) Then
    rsCheck_first = rsCheck_total
  End If
  If (rsCheck_last > rsCheck_total) Then
    rsCheck_last = rsCheck_total
  End If
  If (rsCheck_numRows > rsCheck_total) Then
    rsCheck_numRows = rsCheck_total
  End If
End If
%>



<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsCheck_total = -1) Then

  ' count the total records by iterating through the recordset
  rsCheck_total=0
  While (Not rsCheck.EOF)
    rsCheck_total = rsCheck_total + 1
    rsCheck.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsCheck.CursorType > 0) Then
    rsCheck.MoveFirst
  Else
    rsCheck.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsCheck_numRows < 0 Or rsCheck_numRows > rsCheck_total) Then
    rsCheck_numRows = rsCheck_total
  End If

  ' set the first and last displayed record
  rsCheck_first = 1
  rsCheck_last = rsCheck_first + rsCheck_numRows - 1
  
  If (rsCheck_first > rsCheck_total) Then
    rsCheck_first = rsCheck_total
  End If
  If (rsCheck_last > rsCheck_total) Then
    rsCheck_last = rsCheck_total
  End If

End If
%>



<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
rsCheckDetails_numRows = rsCheckDetails_numRows + Repeat2__numRows
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
<%
if request.form("jewelry") <> "" Then
Recordset1.MoveFirst
qtyarray=split(Request("qty"),",")
for i=0 to ubound(qtyarray)
Recordset1("qty")=trim(qtyarray(i))
Recordset1.update
Recordset1.Movenext
next
jewelry = Request.Querystring("jewelry")
Response.Redirect("inventory_check.asp?jewelry=" & jewelry)
end if
%>
<html>
<head>
<title>Free item inventory</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
   Free item inventory list (<%=(rsCheck_total)%> total) </h5>
  <table class="table table-sm table-hover table-striped">
    <thead class="thead-dark">
    <tr>
      <th>Qty</th>
      <th>Product</th>
    </tr>
  </thead>
    <% i = 0 %>
    <% 
While ((Repeat1__numRows <> 0) AND (NOT rsCheck.EOF)) 
%>
    <% i = i + 1 %>
    <tr>
      <td><%=(rsCheck.Fields.Item("qty").Value)%></td>
      <td><a href="product-edit.asp?ProductID=<%=(rsCheck.Fields.Item("ProductID").Value)%>&info=less""> <%=(rsCheck.Fields.Item("title").Value)%>&nbsp;<%=(rsCheck.Fields.Item("ProductDetail1").Value)%></a></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsCheck.MoveNext()
Wend
%>
  </table>

</div>
</body>
</html><%
rsCheck.Close()
Set rsCheck = Nothing
%>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
