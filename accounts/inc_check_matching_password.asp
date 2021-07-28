<%
	' Get salt from DB by customer ID
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT salt FROM customers WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,request.form("custid")))
	set rsGetSalt = objCmd.Execute()
	
	usersalt = rsGetSalt.Fields.Item("salt").Value
	hashedPass = sha256(usersalt & request.form("password") & extra_key)
	
	set rsGetSalt = nothing
	
	' Check to see if the password matches
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID, password_hashed FROM customers WHERE customer_ID = ? AND password_hashed = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,request.form("custid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("password",200,1,250,hashedPass))
	set rsDuplicatePassFound = objCmd.Execute()
	
	' If a duplicate is found
	if NOT rsDuplicatePassFound.BOF and NOT rsDuplicatePassFound.EOF then
	
		var_matching_password = "yes"
	
	else ' if duplicate not found
	
		var_matching_password = "no"


	end if  ' if duplicate not found


%>