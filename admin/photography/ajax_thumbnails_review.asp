<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
If var_access_level = "Admin" then
	update_column = "thumbnails_approved"
	date_column = "date_thumbs_approved"
else ' if photography
	update_column = "thumbnails_submitted"
	date_column = "date_thumbs_submitted"
end if

if request.form("status") = "approve" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update jewelry set " & update_column & " = 1, " & date_column & " = '" & now() & "' where productID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request.form("productid")))
	objCmd.Execute()

else ' if declined

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update jewelry set thumbnails_submitted = 0 where productID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request.form("productid")))
	objCmd.Execute()

end if

	
	
DataConn.Close()
%>