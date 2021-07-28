<%
' If customer is NOT registered 
if Request.Cookies("ID") = "" then 

guest_extra_key = "zqL)E5fZ8uB%5;yU2~HD" 'Static stored value just for extra security


	
	 ' run if no "cart-userid" cookie is set
	if request.cookies("cartSessionid") = "" then
	
		
		' Generate a random values and store information to the tbl_guest_users table and appropriate cookies
		guest_session = Session.SessionID
		guest_salt = getSalt(32, extraChars)
		hashed_guest_id = sha256(guest_salt & guest_session & guest_extra_key)
		response.cookies("cartSessionid") = Session.SessionID
		response.cookies("cartSelector") = guest_salt
		Response.Cookies("cartSessionid").Expires = DATE + 30
		Response.Cookies("cartSelector").Expires = DATE + 30

		
		'Check that hashed value does not already exist in database before writing value
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_guest_users (guest_id, guest_session_id, guest_salt) VALUES (?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("guest_id",200,1,250,hashed_guest_id))
		objCmd.Parameters.Append(objCmd.CreateParameter("guest_sessionid",200,1,250,guest_session))
		objCmd.Parameters.Append(objCmd.CreateParameter("guest_salt",200,1,250,guest_salt))
		objCmd.Execute()
	
		' Added this line on 6/27/17 to attempt to fix a 500 error    /cart/inc_cart_main.asp Line 26 Incorrect syntax near the keyword 'DEFAULT' 
		var_guest_customer_id = guest_session

	else
	
		'if session cookie is already set then do a compare hash to database to verify before retrieving userID variable

		var_get_cart_session = request.cookies("cartSessionid")
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM tbl_guest_users WHERE guest_session_id = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("guest_session_id",200,1,250,var_get_cart_session))
		set rsGetGuestUser = objCmd.Execute()
		
		if rsGetGuestUser.eof then
			var_guest_customer_id = 987456512
		end if
		while not rsGetGuestUser.eof
			
			hashed_compare_database = sha256(rsGetGuestUser.Fields.Item("guest_salt").Value & rsGetGuestUser.Fields.Item("guest_session_id").Value & guest_extra_key)
	
			hashed_compare_cookies = sha256(request.cookies("cartSelector") & request.cookies("cartSessionid") & guest_extra_key)

			'	response.write "Database hash: " & hashed_compare_database & "<br/><br/>Cookies hash: " & 	hashed_compare_cookies & "<br/><br/>"		
			
			if hashed_compare_database = hashed_compare_cookies then
				var_guest_customer_id = rsGetGuestUser.Fields.Item("guest_unique_id").Value
			'	response.write "Successful match: " & var_guest_customer_id
			end if
		
		rsGetGuestUser.movenext()
		wend
		
	
	end if ' If no cookie is set request.cookies("cartSessionid") = ""
	
end if	' If customer is NOT registered 

' Assign user ID for storing to tbl_carts
if Request.Cookies("ID") = "" then 
	session("guestID") = var_guest_customer_id ' used to transfer cart on signin_transfer page
	var_cart_userid = var_guest_customer_id
	var_db_field = "cart_guest_userid"
else
	var_cart_userid = CustID_Cookie
	var_db_field = "cart_custID"
end if


'response.write "var_cart_userid: " & var_cart_userid & "<br/>" & "var_db_field: " & var_db_field
%>