<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT our_notes FROM sent_items WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,request("ID")))
Set rsGetOldNotes = objCmd.Execute()
%>
<%= rsGetOldNotes.Fields.Item("our_notes").Value %>
<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT tbl_invoice_notes.*, TBL_AdminUsers.name FROM TBL_AdminUsers INNER JOIN tbl_invoice_notes ON TBL_AdminUsers.ID = tbl_invoice_notes.user_id WHERE invoice_id = ? ORDER BY date_created DESC" 
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,request("ID")))
Set rsGetNotes = objCmd.Execute()

if not rsGetNotes.EOF then
%>
<div class="alert alert-success">
<%
while NOT rsGetNotes.EOF %>
		
		<b><%= rsGetNotes.Fields.Item("name").Value %>&nbsp;&nbsp;&nbsp;<%= rsGetNotes.Fields.Item("date_created").Value %></b><br/>
		<%= rsGetNotes.Fields.Item("note").Value %>
		<br/><br/>
<%
rsGetNotes.movenext()
wend
%>
</div>
<%
end if  'if not rsGetNotes.EOF then

DataConn.Close()
%>
</div>