<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM tbl_customer_searches WHERE id = ? and  customer_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,12,request.form("id")))
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,12,CustID_Cookie))
	objCmd.Execute()

%>
{
	"status":"success"
}
<%		

DataConn.Close()
Set DataConn = Nothing
%>
