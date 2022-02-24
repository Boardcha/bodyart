<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("detailid") <> "" then ' check to see if a detailid is provided and if not, just update the main products table	

	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT qty FROM ProductDetails WHERE ProductDetailID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, request.form("detailid") ))
	Set rsGetCurrentQty = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update ProductDetails set qty= ? where ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,request.form("qty")))
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsUpdate = objCmd.Execute()

	'Write info to edits log	
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (?, " & request.form("detailid") & ",'Automated - Updated / Overwrote qty from " & rsGetCurrentQty("qty") & " to " & request.form("qty") & " using the inventory count page for limited bins','" & now() & "')"
	objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
	objCmd.Execute()
	Set objCmd = Nothing

end if

DataConn.Close()
Set rsResearch = Nothing
%>