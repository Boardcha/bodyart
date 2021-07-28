<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT * FROM dbo.QRYTopWaitingListItems WHERE qty >= waiting_qty ORDER BY title ASC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
rsGetRecords_numRows = 0

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

<link href="../CSS/Admin.css" rel="stylesheet" type="text/css" />
<title>Waiting list items</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="content-grey admin-content">

    <table class="admin-table">
		<thead>
		<tr>
		<th colspan="5">
		Waiting list review (<%=(rsGetRecords_total)%>)
		</th>
		</tr>
		<tr>
    <th>Waiting</th>
    <th>Email</th>
		<th>Item</th>
		<th>Detail</th>
		<th>Stock level</th>
		</tr>
		</thead>
      <% 
While NOT rsGetRecords.EOF
%>
<tr id="<%= rsGetRecords.Fields.Item("ID").Value %>">
          <td><%=(rsGetRecords.Fields.Item("howmany").Value)%></td>
          <td><i class="fa fa-times-circle fa-lg text-danger mr-4 btn btn-delete" data-id="<%= rsGetRecords.Fields.Item("ID").Value %>"></i><%=(rsGetRecords.Fields.Item("email").Value)%></td>
      <td><a href="product-edit.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%>"><%=(rsGetRecords.Fields.Item("title").Value)%></a></td>
      <td><%=(rsGetRecords.Fields.Item("ProductDetail1").Value)%> <%=(rsGetRecords.Fields.Item("ProductDetail1").Value)%></td>
      <td>IN STOCK: <%=(rsGetRecords.Fields.Item("qty").Value)%> </td>
</tr>
      <% 
  rsGetRecords.MoveNext()
Wend
%>
    </table>
		
</div>
</body>
</html>
<%
rsGetRecords.Close()
Set rsGetRecords = Nothing
%>
<script type="text/javascript">
  // Delete waiting list item
  $(document).on("click", ".btn-delete", function(event){
      var id = $(this).attr("data-id");

      $.ajax({
      method: "POST",
      url: "/admin/inventory/ajax-delete-waiting-list-item.asp",
      data: {id: id}
      })
      .done(function(msg ) {
          $('#' + id).addClass('table-danger');
          $('#' + id).fadeOut('slow');
      })
      .fail(function(msg) {
          alert('FAILED');
      });
  });
</script>
