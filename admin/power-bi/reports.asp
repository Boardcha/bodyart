<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set Cmd_rsGetTotalOrder = Server.CreateObject("ADODB.command")
Cmd_rsGetTotalOrder.ActiveConnection = DataConn
Cmd_rsGetTotalOrder.CommandText = "SELECT Count(*) AS Total_ToShip FROM sent_items  WHERE ship_code = 'paid' AND (shipped = 'Pending shipment' OR shipped = 'SHIPPING BACKORDER' OR shipped = 'RETURN ENVELOPE' OR shipped = 'RESHIP PACKAGE')"
Set rsGetTotalOrder = Cmd_rsGetTotalOrder.Execute()
%>
<title><%= request.querystring("reportName") %> - Report</title>

<html>
    <body>
    <!--#include virtual="/admin/admin_header.asp"-->
<style>
    a {color:black}
</style>
    <div class="p-2">
    <h4 class="mb-1"><%= request.querystring("reportName") %></h4>

    <iframe class="" id="load-report" width="1600px" height="6000px" frameborder="0" allowFullScreen="true" scrolling="no" src="https://app.powerbi.com/reportEmbed?reportId=<%= request.querystring("reportId") %>&pageName=<%= request.querystring("pageName") %>&reportName=<%= request.querystring("reportName") %>&navContentPaneEnabled=false&filterPaneEnabled=false&autoAuth=true&ctid=06bc9374-9044-4ccb-8d1c-84eb80fc2e89&config=eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly93YWJpLXVzLWNlbnRyYWwtYS1wcmltYXJ5LXJlZGlyZWN0LmFuYWx5c2lzLndpbmRvd3MubmV0LyJ9"></iframe>

    

        </div><!-- body padding-->
    </body>
</div>



</html>
