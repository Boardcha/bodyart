<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT ID, DetailID, name, email, title, ProductDetail1, waiting_qty, date_added FROM dbo.QRYWaitingList WHERE DetailID = ? ORDER BY date_added ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("DetailID",3,1,20, request.querystring("DetailID")  ))

set rsShowWaitingList = Server.CreateObject("ADODB.Recordset")
rsShowWaitingList.CursorLocation = 3 'adUseClient
rsShowWaitingList.Open objCmd
var_total_waiting = rsShowWaitingList.RecordCount
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>View waiting list</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>

  <!--#include file="admin_header.asp"-->
  <div class="p-2">
  <% if rsShowWaitingList.EOF then %>
  <h4>Nobody currently on the waiting list</h4>
<% else %>
 <h4><%=(rsShowWaitingList.Fields.Item("title").Value)%>&nbsp;<%=(rsShowWaitingList.Fields.Item("ProductDetail1").Value)%> (<%= var_total_waiting %>)   </h4>
 
<table class="table table-sm table-striped table-borderless w-50">
  <thead class="thead-dark">
<tr> 
  <th>Email</th>
<th>Qty wanted</th>
<th>Date added</th>
</tr>
  </thead>

    <% 
While NOT rsShowWaitingList.EOF
%>
    <tr id="row-<%= rsShowWaitingList("ID") %>">
        <td><i class="btn btn-sm btn-danger fa fa-trash-alt mr-5 delete-row"  data-id="<%= rsShowWaitingList("ID") %>"></i><%=(rsShowWaitingList.Fields.Item("email").Value)%></td>
        <td><%= rsShowWaitingList("waiting_qty") %></td>
        <td><%= rsShowWaitingList("date_added") %></td>
    </tr>
      <% 
  rsShowWaitingList.MoveNext()
Wend
%>
  </table>
  <% end if %>
</div>

</body>
</html>
<%
rsShowWaitingList.Close()
Set rsShowWaitingList = Nothing
%>
<script type="text/javascript">
// Delete item from waiting list
$(".delete-row").click(function(){ 
  var id = $(this).attr("data-id");

  $.ajax({
  method: "POST",
  url: "inventory/ajax-delete-waiting-list.asp",
  data: {id: id}
  })
  .done(function( msg ) {
    $('#row-' + id).fadeOut("slow");
  })
  .fail(function(msg) {
    alert("DELETE FAILED " + id);
  });
  
});
</script>