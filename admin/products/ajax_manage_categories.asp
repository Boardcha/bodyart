<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.QueryString("addCategory")="yes"	then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT category_name, category_tag FROM TBL_Categories WHERE category_tag=?"
	objCmd.Parameters.Append(objCmd.CreateParameter("category_tag", 200, 1, 100, request.QueryString("category_tag")))
	Set rsCategories = objCmd.Execute()
	'If category does not exist
	If rsCategories.EOF Then
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_Categories(category_name, category_tag) VALUES (?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("category_name", 200, 1, 100, request.QueryString("category_name")))
		objCmd.Parameters.Append(objCmd.CreateParameter("category_tag", 200, 1, 100, request.QueryString("category_tag")))
		objCmd.Execute()
	End If
	
elseif request.QueryString("deleteCategory")="yes" then

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_Categories WHERE category_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("category_id", 3, 1, 15, request.QueryString("category_id")))
	objCmd.Execute()
	
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT category_name, category_tag FROM TBL_Categories ORDER BY category_name ASC"
Set rsGetCategories = objCmd.Execute()

	
While NOT rsGetCategories.EOF 
	Response.Write("<option value=""" & rsGetCategories.Fields.Item("category_tag").Value & """>" & rsGetCategories.Fields.Item("category_name").Value & "</option>")
	rsGetCategories.MoveNext()
Wend

Set rsGetCategories = Nothing
Set rsCategories = Nothing
DataConn.Close()
%>