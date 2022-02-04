<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_customer_searches (search_url, customer_id, date_added) VALUES (?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("url",200,1,500,request.form("url") & ","))
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,12,CustID_Cookie))
	objCmd.Parameters.Append(objCmd.CreateParameter("date_added",200,1,75, now()))
	objCmd.Execute()

%>
{
	"status":"success"
}
<%		

DataConn.Close()
Set DataConn = Nothing
%>
