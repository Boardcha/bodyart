<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
{
<%
invoiceid = request.form("id")

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ? AND ship_code = 'paid' AND (shipped = 'Pending...' OR shipped = 'Pending shipment' OR shipped = 'Review' OR shipped = 'CUSTOM ORDER IN REVIEW') OR shipped = 'CUSTOM COLOR IN PROGRESS')"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,invoiceid))
	Set rsGetOrder = objCmd.Execute()

if not rsGetOrder.eof then	

if CLng(CustID_Cookie) = CLng(rsGetOrder.Fields.Item("customer_ID").Value) then
%>
    "status":"success"
<% else ' ' CustID_Cookie = order customerID
%>
    "status":"fail"
<%
	end if ' CustID_Cookie = order customerID
else '==== order no longer eligible to be cancelled
%>
    "status":"fail"
<%
end if ' not rsGetOrder.eof
DataConn.Close()
%>
}