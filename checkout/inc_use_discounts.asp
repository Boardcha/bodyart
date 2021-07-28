<%
' if a one time use coupon was applied, then mark it as being used. This will actually mark ANY coupon as used, but it doesn't really matter if it can be used more than once.
if Session("CouponCode") <> "" then
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBLDiscounts SET coupon__single_redeemed = 1, coupon_single_invoice = ? WHERE discountcode = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,Session("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,30,Session("CouponCode")))
	objCmd.Execute()
end if

' if gift cert code applied is found
If Session("GiftCertAmount") > 0 then

	' Write to the credits table what it was used on
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_Credits_UsedOn (OriginalCreditID, InvoiceUsedOn) VALUES (?, ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("CertificateID",3,1,10,Session("GiftCertID")))
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,Session("invoiceid")))
	objCmd.Execute()

	'Get certificate by code user enters in
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID, code, invoice, amount FROM TBLcredits WHERE ID = ? and amount > 0"
	objCmd.Parameters.Append(objCmd.CreateParameter("giftcertid",3,1,30,Session("GiftCertID")))
	Set rsGetCert = objCmd.Execute()
	
	' if certificate is found 
	if Not rsGetCert.EOF then

	
	' if there is an amount that needs to be charged then write the gift cert as $0 balance to the database
	If var_total_giftcert_dueback <= 0 then

	'Bug testing ----------
	'response.write "<br/>" & Session("GiftCertAmount") & "<br/>"
	'response.write Session("GiftCertID") & "<br/>"	
	
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBLcredits SET amount = 0 WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("GiftCertID",3,1,30,Session("GiftCertID")))
		objCmd.Execute()

	else ' if the customer owes $0 AND there's a balance due back to the certificate then write that back to the database

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBLcredits SET amount = ? WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("amount",6,1,10,FormatNumber(var_total_giftcert_dueback, -1, -2, -2, -2)))
		objCmd.Parameters.Append(objCmd.CreateParameter("GiftCertID",3,1,30,Session("GiftCertID")))
		objCmd.Execute()
	%>
	
	,"covered_infull_giftcert":"yes"

	<%
	end if  ' write balance due to database

	end if ' if certificate is found in database

' if gift cert code applied is found
End if	
%>