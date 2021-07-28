<%
' If customer is NOT registered --------------------------------
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

end if	


	'Remove item from cart
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM tbl_carts WHERE cart_id = ? AND " & var_db_field & " = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10,request.form("cart_id")))
	objCmd.Parameters.Append(objCmd.CreateParameter("cust_id",3,1,10,var_cart_userid))
	objCmd.Execute()
%>