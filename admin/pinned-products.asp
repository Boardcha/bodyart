<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM jewelry WHERE pinned_product = 1"
Set rsGetPinnedItems = objCmd.Execute()

%>

<html>
<head>
<title>Pinned Products</title>
</head>
<body>

<!--#include file="admin_header.asp"-->
<div class="m-3">
<% If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE %>

<h5>Pinned Products</h5>

<table class="table table-striped table-sm table-hover">
    <thead class="thead-dark">
            <tr>
                    <th>
                      Title
                  </th>
                  </tr>
    </thead>
    <tbody>
<%
While NOT rsGetPinnedItems.EOF 
%>
        <tr>
            <td><a href="product-edit.asp?ProductID=<%= rsGetPinnedItems.Fields.Item("ProductID").Value %>"><%= rsGetPinnedItems.Fields.Item("title").Value %></a></td>
        </tr>

<% 
rsGetPinnedItems.MoveNext()
Wend
%>
    </tbody>  
</table>
<%
else ' unathorized access error %>
Not accessible
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

</div>
</body>
</html>
<%
DataConn.Close()
Set rsGetPinnedItems = Nothing
%>
