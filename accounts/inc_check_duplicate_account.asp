<%
if Request.Form("e-mail") <> "" then
	var_email = Request.Form("e-mail")
end if
if Request.Form("email") <> "" then
	var_email = Request.Form("email")
end if

if var_email <> "" then
	' Check to see if account by e-mail already exists
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID, email FROM customers WHERE email = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50,var_email))
	set rsDuplicateFound = objCmd.Execute()
	
	' If a duplicate is found
	if NOT rsDuplicateFound.BOF and NOT rsDuplicateFound.EOF then
	
		var_duplicate_account = "yes"
	
	else ' if duplicate not found
	
		var_duplicate_account = "no"


	end if  ' if duplicate not found
end if ' if an email is found at all
%>