<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% response.Buffer=false
Server.ScriptTimeout=300
 %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/emails/function-send-email.asp"-->
<%
Dim rsGetCustomers
Dim rsGetCustomers_cmd
Dim rsGetCustomers_numRows

Set rsGetCustomers_cmd = Server.CreateObject ("ADODB.Command")
rsGetCustomers_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCustomers_cmd.CommandText = "SELECT * FROM dbo.QRY_WaitingList_Notify WHERE qty >= waiting_qty AND customer_notified = 0" 
rsGetCustomers_cmd.Prepared = true

Set rsGetCustomers = rsGetCustomers_cmd.Execute

mailer_type = "notify waiting list"
While NOT rsGetCustomers.EOF 
	if rsGetCustomers.Fields.Item("email").Value <> "" then
		%>
		<!--#include virtual="/emails/email_variables.asp"-->
		<%
	end if

	set Command1 = Server.CreateObject("ADODB.Command")'create command object
	Command1.ActiveConnection = MM_bodyartforms_sql_STRING 'connection string
	Command1.CommandText = "UPDATE TBLWaitingList SET customer_notified = 1 WHERE ID = " & rsGetCustomers.Fields.Item("ID").Value
	Command1.Execute() 
	Set Command1 = Nothing
	rsGetCustomers.MoveNext()
Wend

'====== GET NEW WAITING LIST TOTAL
set cmd_rsGetWaitingList = Server.CreateObject("ADODB.command")
cmd_rsGetWaitingList.ActiveConnection = DataConn
cmd_rsGetWaitingList.CommandText = "SELECT Count(*) AS Total_Waiting FROM dbo.QRYTopWaitingListItems WHERE qty >= waiting_qty"
Set rsGetWaitingList = cmd_rsGetWaitingList.Execute()
%>
{
  "total":"<%= rsGetWaitingList.Fields.Item("Total_Waiting").Value %>"
}
<%
rsGetCustomers.Close()
Set rsGetCustomers = Nothing
Set rsGetCustomers_cmd = Nothing
%>
