<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
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

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE tbl_carts SET cart_detailId = ? WHERE cart_id = ? AND " & var_db_field & " = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("cart_detailId",3,1,10, request.form("detailid")))
objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10, request.form("cartid")))
objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, var_cart_userid))
objCmd.Execute()


DataConn.Close()
Set DataConn = Nothing
Set rs_getCart = Nothing
%>