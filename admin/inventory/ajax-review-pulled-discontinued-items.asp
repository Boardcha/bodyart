<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	column_title = request.form("column")
	
	' protection to stop page from sql injection
%>
<!--#include file="../../functions/sql_injection_prevention.asp" -->
<%
	
	column_value = request.form("value")
	id = request.form("id")
	friendly_name = request.form("friendly_name")
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET qty = ? where ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,15, column_value ))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, id))
	objCmd.Execute()

	' ====== INSERT EDITS LOG WITH ALL INFORMATION 
	Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, edit_date, detail_id, description) VALUES(?, GETDATE(), ?, 'Automated message: Manually updated qty to ' + ? + ' from admin reviewing pulled discontinued items')"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("detail_id",3,1,15, id ))
    objCmd.Parameters.Append(objCmd.CreateParameter("qty",200,1,10, column_value ))
    objCmd.Execute 
		

Set rsGetUser = nothing
DataConn.Close()
%>