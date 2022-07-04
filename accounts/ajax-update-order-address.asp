<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
For Each item In Request.Form
    'Response.Write "Key: " & item & " - Value: " & Request.Form(item) & "<BR />"
Next

invoiceid = request.form("id")

    set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ? AND ship_code = 'paid'"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,invoiceid))
	Set rsGetOrder = objCmd.Execute()

if CLng(CustID_Cookie) = CLng(rsGetOrder.Fields.Item("customer_ID").Value) then

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE sent_items SET customer_first = ?, customer_last = ?, address = ?, address2 = ?, city = ?, state = ?, province = ?, zip = ? WHERE id = ?"

    objCmd.Parameters.Append(objCmd.CreateParameter("@First",200,1,30, request.form("first") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("@Last",200,1,30, request.form("last") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("@Street",200,1,75, request.form("shipping-address") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("@Address2",200,1,75, request.form("shipping-address2") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("@City",200,1,50, request.form("shipping-city") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("@State",200,1,50, request.form("shipping-state") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("@Province",200,1,30, request.form("shipping-province-canada") & "" & request.form("shipping-province") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("@Zip",200,1,15, request.form("shipping-zip") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10, invoiceid))
    objCmd.Execute()

    ' -------------  Insert notes about order
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,1))
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,invoiceid))
    objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,"Automated message - Customer updated address through website"))
    objCmd.Execute()
%>
{
    "status":"success",
    "invoice": "<%= invoiceid %>"
}
<%
else ' ' CustID_Cookie = order customerID
%>
{
	"status":"fail"
}
<%	
	end if ' CustID_Cookie = order customerID
DataConn.Close()
%>
