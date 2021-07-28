<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("orderdetailid") <> "" and  var_access_level = "Admin" OR var_access_level = "Manager" then

  '===== GET ALL ERRORS ====================================
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_OrderSummary SET item_problem = 0 WHERE OrderDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,15, request.form("orderdetailid") ))
    objCmd.Execute()

end if
%>