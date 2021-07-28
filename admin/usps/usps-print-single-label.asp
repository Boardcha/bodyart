<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<style>
 /* style sheet for "letter" printing */
    @page {
        size: 4in 6in;
        margin: 0
    }
</style>
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT usps_base64_shipping_label FROM sent_items WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12,request.querystring("invoiceid")))
Set rsGetOrders = objCmd.Execute()

While NOT rsGetOrders.EOF
shipping_label = rsGetOrders("usps_base64_shipping_label")
%>

<%
'===== DETECT WHETHER INTERNATIONAL LABEL ======
If instr(shipping_label, "R0lGOD") > 0 then

	split_img = Split(rsGetOrders("usps_base64_shipping_label"), "R0lGOD")
%>
	<img src="data:image/gif;base64, R0lGOD<%= split_img(1) %>" style="width:4in;height:6in"> 
<% else %>
	<img src="data:image/png;base64, <%= shipping_label %>" style="width:4in;height:6in">
<% end if %>
<div style="page-break-after: always"></div>

<%
rsGetOrders.MoveNext()
Wend
%>
