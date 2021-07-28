<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("productid") <> 0 and request.form("active") = "yes" then 
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET new_page_date = '" & now() & "' WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,request.form("productid")))
	objCmd.Execute()

else

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET new_page_date = '1/1/2000' WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,request.form("productid")))
	objCmd.Execute()

end if
DataConn.Close()
%>
{  
   "status":"success"
}