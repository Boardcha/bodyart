<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
Dim rsGetRestockItems__MMColParam
rsGetRestockItems__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsGetRestockItems__MMColParam = Request.QueryString("ID")
End If
%>
<%
SortBy = Request.Querystring("SortBy")

Dim rsGetRestockItems
Dim rsGetRestockItems_cmd
Dim rsGetRestockItems_numRows

Set rsGetRestockItems_cmd = Server.CreateObject ("ADODB.Command")
rsGetRestockItems_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRestockItems_cmd.CommandText = "SELECT * FROM dbo.QRY_PORestock WHERE PurchaseOrderID = ? ORDER BY " & SortBy 
rsGetRestockItems_cmd.Prepared = true
rsGetRestockItems_cmd.Parameters.Append rsGetRestockItems_cmd.CreateParameter("param1", 5, 1, -1, rsGetRestockItems__MMColParam) ' adDouble

Set rsGetRestockItems = rsGetRestockItems_cmd.Execute
rsGetRestockItems_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetRestockItems_numRows = rsGetRestockItems_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsGetRestockItems_total
Dim rsGetRestockItems_first
Dim rsGetRestockItems_last

' set the record count
rsGetRestockItems_total = rsGetRestockItems.RecordCount

' set the number of rows displayed on this page
If (rsGetRestockItems_numRows < 0) Then
  rsGetRestockItems_numRows = rsGetRestockItems_total
Elseif (rsGetRestockItems_numRows = 0) Then
  rsGetRestockItems_numRows = 1
End If

' set the first and last displayed record
rsGetRestockItems_first = 1
rsGetRestockItems_last  = rsGetRestockItems_first + rsGetRestockItems_numRows - 1

' if we have the correct record count, check the other stats
If (rsGetRestockItems_total <> -1) Then
  If (rsGetRestockItems_first > rsGetRestockItems_total) Then
    rsGetRestockItems_first = rsGetRestockItems_total
  End If
  If (rsGetRestockItems_last > rsGetRestockItems_total) Then
    rsGetRestockItems_last = rsGetRestockItems_total
  End If
  If (rsGetRestockItems_numRows > rsGetRestockItems_total) Then
    rsGetRestockItems_numRows = rsGetRestockItems_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsGetRestockItems_total = -1) Then

  ' count the total records by iterating through the recordset
  rsGetRestockItems_total=0
  While (Not rsGetRestockItems.EOF)
    rsGetRestockItems_total = rsGetRestockItems_total + 1
    rsGetRestockItems.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsGetRestockItems.CursorType > 0) Then
    rsGetRestockItems.MoveFirst
  Else
    rsGetRestockItems.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsGetRestockItems_numRows < 0 Or rsGetRestockItems_numRows > rsGetRestockItems_total) Then
    rsGetRestockItems_numRows = rsGetRestockItems_total
  End If

  ' set the first and last displayed record
  rsGetRestockItems_first = 1
  rsGetRestockItems_last = rsGetRestockItems_first + rsGetRestockItems_numRows - 1
  
  If (rsGetRestockItems_first > rsGetRestockItems_total) Then
    rsGetRestockItems_first = rsGetRestockItems_total
  End If
  If (rsGetRestockItems_last > rsGetRestockItems_total) Then
    rsGetRestockItems_last = rsGetRestockItems_total
  End If

End If
%>
<html>
<head>

<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Put items in stock</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#FFFFFF">
<a href="PurchaseOrders_Print.asp?ID=<%= Request.Querystring("ID") %>&SortBy=ProductDetailID ASC" class="Link_ItemDetails">Sort by #</a>&nbsp; |&nbsp;&nbsp;<a href="PurchaseOrders_Print.asp?ID=<%= Request.Querystring("ID") %>&SortBy=title ASC" class="Link_ItemDetails">Sort by name</a> <p> 
 <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetRestockItems.EOF)) 
%>
<span class="smallestfont"><strong><font size="+1"><img src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetRestockItems.Fields.Item("picture").Value)%>" alt="Image" width="50" height="50" align="absmiddle">&nbsp;&nbsp;<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%></font></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Purchased: <%=(rsGetRestockItems.Fields.Item("POAmount").Value)%> &nbsp;&nbsp; <%=(rsGetRestockItems.Fields.Item("title").Value)%><a href="product-edit.asp?ProductID=<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>&info=less" target="_blank" class="EditSelect_Links">&nbsp;</a><%=(rsGetRestockItems.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("ProductDetail1").Value)%></span> <br>
<hr width="100%" size="1">
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetRestockItems.MoveNext()
Wend
%>

<p><br>
</p>
</body>
</html>
<%
rsGetRestockItems.Close()
Set rsGetRestockItems = Nothing
%>
