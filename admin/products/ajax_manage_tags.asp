<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.QueryString("addTag")="yes"	then

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_Product_Tags(tag) VALUES (?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("tag", 200, 1, 100, request.QueryString("tag")))
	objCmd.Execute()
	
elseif request.QueryString("deleteTag")="yes" then

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_Product_Tags WHERE tagID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("tagID", 3, 1, 15, request.QueryString("tagID")))
	objCmd.Execute()
	
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT tag FROM TBL_Product_Tags ORDER BY tag ASC"
Set rs_getTags = objCmd.Execute()

	
While NOT rs_getTags.EOF 
	Response.Write("<option value=""" & rs_getTags.Fields.Item("tag").Value & """>" & rs_getTags.Fields.Item("tag").Value & "</option>")
	rs_getTags.MoveNext()
Wend

Set rs_getTags = Nothing
DataConn.Close()
%>