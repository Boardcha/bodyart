<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT DetailID, name, email, title, ProductDetail1 FROM dbo.QRYWaitingList WHERE DetailID = ? ORDER BY name ASC"
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

    <% 
While NOT rsShowWaitingList.EOF
%>
<tr>
        <td><%=(rsShowWaitingList.Fields.Item("name").Value)%></td>
        <td><%=(rsShowWaitingList.Fields.Item("email").Value)%></td>
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
