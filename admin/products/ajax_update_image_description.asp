<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE tbl_images SET img_description = ? WHERE img_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("img_description",200,1,50,request.form("img_description")))
	objCmd.Parameters.Append(objCmd.CreateParameter("img-id",3,1,15,request.form("imgid")))
	objCmd.Execute()

DataConn.Close()
%>