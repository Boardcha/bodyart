<%
' START USE STORE CREDIT - This is only included where the payment has been approved
' =================================================================
if session("storeCredit_amount") <> 0 and CustID_Cookie <> "" and CustID_Cookie <> 0 then

	if var_grandtotal > 0 and session("storeCredit_amount") > var_grandtotal then ' all used up, 0 balance
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET credits = 0 WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		objCmd.Execute()
%>
 ,"store_credit":"used_entire"
<%
	else
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET credits = ? WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("grand_credit",6,1,10,FormatNumber(var_credit_due_todb, -1, -2, -2, -2)))
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		objCmd.Execute()

' Do not display this Json on last page of Paypal checkout. This should only be used for credit card checkout for ajax information.
if request.querystring("token") = "" then
%>
,"store_credit":"used_partial"
<%	
end if	' token is empty
	end if

end if
' END USE STORE CREDIT  =========++========================================================


' START USE NOW free gift credits  ========================================================
if session("credit_now") <> "" then

	' Nothing needs to be done for free gift credits that are used on current order
	'	FormatNumber(session("credit_now"), -1, -2, -2, -2)

end if
' END USE NOW free gift credits  =========++===============================================

' START SAVE FOR LATER free gift credits  ========================================================
if session("credit_later") <> "" and CustID_Cookie <> "" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET credits = credits + ? WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("credit_later",6,1,10,FormatNumber(session("credit_later"), -1, -2, -2, -2)))
	objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
	objCmd.Execute()

end if
' END SAVE FOR LATER free gift credits  =========++===============================================
%>