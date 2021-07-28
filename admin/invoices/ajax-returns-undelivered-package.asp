<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="emails/function-send-email.asp"-->

<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, request.form("invoice")))
Set rsGetOrder = objCmd.Execute()

var_reason = request.form("reason")

if rsGetOrder.Fields.Item("company").Value <> "" then
    var_company = rsGetOrder.Fields.Item("company").Value & "<br>"
end if
if rsGetOrder.Fields.Item("address2").Value <> "" then
    var_address2 = rsGetOrder.Fields.Item("address2").Value & "<br>"
end if

' ---------- Set order to Package came back status
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE sent_items SET shipped = 'PACKAGE CAME BACK' WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, request.form("invoice")))
objCmd.Execute()

' ---------- Add a note to the order
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, request.form("invoice")))
objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250, "Package came back. Reason: " &  var_reason))
objCmd.Execute()

mailer_type = "returns"
%>
<!--#include virtual="emails/email_variables.asp"-->
<%

DataConn.Close()
Set rsGetOrder = Nothing
%>
