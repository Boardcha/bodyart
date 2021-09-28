<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("orderdetailid") <> "" then
	If request.form("checked") = "true" Then isChecked = 1 Else isChecked =0
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_OrderSummary SET problem_reviewed = ? WHERE OrderDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("problem_reviewed",3,1,15, isChecked ))
    objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,15, request.form("orderdetailid") ))
    objCmd.Execute()

end if
%>