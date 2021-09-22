<%@ Language=VBScript %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../functions/token.asp"-->
<!--#include file="../../functions/salt.asp"-->
<!--#include file="../../functions/hash_extra_key.asp"-->
<%
' UPDATE PASSWORD -------------------------------------------
' Generate new salt and re-hash new password
if request.form("password") <> "" then

	salt = getSalt(32, extraChars)
	newPass = sha256(salt & request.form("password") & extra_key)
	token = sha256(request.form("token"))

	' DOES TOKEN EXIST? ----------------------------------------
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_AdminUsers WHERE reset_token = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("token",200,1,250,token))
	Set rsTokenExists = objCmd.Execute()
	
	if NOT rsTokenExists.BOF and NOT rsTokenExists.EOF then

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_AdminUsers SET pass_last_updated = '" & now() & "',  password_hashed = ?, salt = ?, reset_token = '', reset_token_timestamp = '' WHERE reset_token = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("pass",200,1,250,newPass))
		objCmd.Parameters.Append(objCmd.CreateParameter("salt",200,1,250,salt))
		objCmd.Parameters.Append(objCmd.CreateParameter("token",200,1,250,token))
		objCmd.Execute()
%>
		{  
		   "status":"success"
		}
<%
	
	else ' empty
%>
		{  
		   "status":"fail"
		}
<%	
	end if
	
	set rsTokenExists = Nothing

end if ' if ID found
' UPDATE PASSWORD -------------------------------------------


' REQUEST A RESET TOKEN -------------------------------------
if request.form("id") <> "" then

	mailer_type = "admin_reset_user_password"
	reset_token = getToken(40, extraChars)
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_AdminUsers SET reset_token = '" & sha256(reset_token) & "', reset_token_timestamp = '" & now() & "' where ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,request.form("id")))
	objCmd.Execute()
%>
<!--#include file="../../emails/function-send-email.asp"-->
<!--#include file="../../emails/email_variables.asp"-->
<%
end if ' if ID found
' REQUEST A RESET TOKEN -------------------------------------

DataConn.Close()
Set DataConn = Nothing
%>

