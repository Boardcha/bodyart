<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->

<!--#include virtual="/functions/encrypt.asp"-->
<!--#include virtual="/functions/token.asp"-->
<!--#include virtual="/functions/hash_extra_key.asp"-->
<%
' Get salt from DB by customer ID
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT salt, password_hashed, registered_with_social_login FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,CustID_Cookie))
set rsUserInfo = objCmd.Execute()

usersalt = rsUserInfo.Fields.Item("salt").Value
hashedPass = sha256(usersalt & request.form("current_password") & extra_key)

' Check to see if the password matches
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT customer_ID, password_hashed FROM customers WHERE customer_ID = ? AND password_hashed = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
objCmd.Parameters.Append(objCmd.CreateParameter("password",200,1,250,hashedPass))
set rsMatchingPassFound = objCmd.Execute()

' If a match is found
if (NOT rsMatchingPassFound.BOF and NOT rsMatchingPassFound.EOF) OR rsUserInfo("registered_with_social_login") = true then
	'If they registered with a social plugin, once they set a password on account-profile.asp we need to set registered_with_social_login = 0 in DB
	'Now They have regular password and If they want to change their password, their current password should be asked
	
	' Re-hash new password to save in BAF database
	salt = getSalt(32, extraChars)
	newPass = sha256(salt & request.form("password") & extra_key)

'	response.write "salt" & salt & " /  new pass: " & newPass & " password: " & request.form("password")

	
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET pass_last_updated = '" & now() & "',  password_hashed = ?, salt = ?, reset_token = '', reset_token_timestamp = '', registered_with_social_login = 0  WHERE customer_ID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("password",200,1,250,newPass))
		objCmd.Parameters.Append(objCmd.CreateParameter("salt",200,1,250,salt))
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,CustID_Cookie))
		objCmd.Execute()

%>
{
	"status":"success"
}
<%


else ' if match not found

%>
{
	"status":"fail"
}
<%
end if  ' if match not found


DataConn.Close()
Set DataConn = Nothing
%>