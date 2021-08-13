<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.QueryString("addMaterial")="yes"	then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_Materials WHERE material_name=?"
	objCmd.Parameters.Append(objCmd.CreateParameter("material", 200, 1, 100, request.QueryString("material")))
	Set rsMaterials = objCmd.Execute()
	'If material does not exist
	If rsMaterials.EOF Then
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_Materials (material_name, toggle_wearable) VALUES (?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("material", 200, 1, 100, request.QueryString("material")))
		objCmd.Parameters.Append(objCmd.CreateParameter("toggle_wearable", 11, 1, 1, request.QueryString("iswearable")))
		objCmd.Execute()
	End If
	
elseif request.QueryString("deleteMaterial")="yes" then

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_Materials WHERE material_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("material_id", 3, 1, 15, request.QueryString("material_id")))
	objCmd.Execute()
	
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Materials ORDER BY material_name ASC"
Set rs_getMaterials = objCmd.Execute()

	
While NOT rs_getMaterials.EOF 
	materials = materials & "<option value='" & rs_getMaterials.Fields.Item("material_name").Value & "'>" & rs_getMaterials.Fields.Item("material_name").Value & "</option>"
	If rs_getMaterials("toggle_wearable") = true Then wearable_materials = wearable_materials & "<option value='" & rs_getMaterials.Fields.Item("material_name").Value & "'>" & rs_getMaterials.Fields.Item("material_name").Value & "</option>"
	rs_getMaterials.MoveNext()
Wend

Set rs_getMaterials = Nothing
Set rsMaterials = Nothing
DataConn.Close()
%>
{  
   "materials":"<%= materials %>",
   "wearable_materials":"<%= wearable_materials %>"
}