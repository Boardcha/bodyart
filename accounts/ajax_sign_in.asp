<% @LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/functions/hash_extra_key.asp"-->
<!--#include virtual="/functions/encrypt.asp"-->
<%
' Check to see if info matches and a user is found

	' Get salt from DB by customer ID
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT salt FROM customers WHERE email = ? ORDER BY customer_ID ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,Request.Form("email")))
	set rsGetSalt = objCmd.Execute()
	
	If Not rsGetSalt.EOF Or Not rsGetSalt.BOF Then
	
		usersalt = rsGetSalt.Fields.Item("salt").Value
	
	End if
	
		hashed_pass = sha256(usersalt & Request.Form("password") & extra_key)
	

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID, password_hashed, salt, active FROM customers WHERE email = ? AND password_hashed = ? ORDER BY customer_ID ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,Request.Form("email")))
	objCmd.Parameters.Append(objCmd.CreateParameter("password",200,1,250,hashed_pass))
	set rsGetUser = objCmd.Execute()




	If rsGetUser.eof then 
	
		session("login_email") = Request.Form("email")
		%>
		{ "status":"logged-out" }
		<% 
	else
		if rsGetUser("active")= true then
			%>
			{ "status":"logged-in" }	
			<%
		
	' Write last login date
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET last_login = '" & now() & "' WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,rsGetUser.Fields.Item("customer_ID").Value))
	objCmd.Execute()
	


	' Set session variable to modify shipping/billing information in account
	session("custID_account") = rsGetUser.Fields.Item("customer_ID").Value
	var_our_custid = rsGetUser.Fields.Item("customer_ID").Value

	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

	password = "3uBRUbrat77V"
	data = rsGetUser.Fields.Item("customer_ID").Value

	encrypted = objCrypt.Encrypt(password, data)
	Response.Cookies("ID") = encrypted
	Response.Cookies("ID").Expires = DATE + 30


	Set objCrypt = Nothing
%>
<!--#include virtual="/accounts/inc_transfer_cart_contents.asp" -->
<%
	if Request.Cookies("ID") <> "" then 

		' decrypt customer ID cookie
		Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

		password = "3uBRUbrat77V"
		data = request.Cookies("ID")

		If len(data) > 5 then ' if
			decrypted = objCrypt.Decrypt(password, data)
		end if

		  if data <> decrypted then
			  CustID_Cookie = decrypted
		  else
			  CustID_Cookie = 0
		  end if

		Set objCrypt = Nothing

	
	end if ' if user is logged in 

	else '==== IF USER HAS NOT CLICKED ACTIVATION EMAIL LINK
	%>
		{ "status":"not-active" }			
	<%end if
	
End If ' end rsGetUser.EOF
	
DataConn.Close()
Set DataConn = Nothing
%>