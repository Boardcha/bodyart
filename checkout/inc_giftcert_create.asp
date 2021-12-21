<% 
' If a gift certificate is in the cart ... set on page cart/inc_cart_loopitems-begin.asp
If var_giftcert = "yes"  Then

		var_cert_code = strRandomCode
	
		' Call function
		var_cert_code = CheckDupe(var_cert_code)

		' Set array
		certificate_array =split(rs_getCart.Fields.Item("cart_preorderNotes").Value,"{}")

		' Set variables
		done_mailing_certs = "no" ' send out an email variable
		gift_amount = FormatNumber(var_lineTotal, -1, -2, -2, -2)
		message = certificate_array(2)
		your_name = certificate_array(1)
		rec_name = certificate_array(3)
		rec_email = certificate_array(0)
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBLCredits (invoice, name, rec_name, rec_email, message, code, amount, certificate_original_amount) VALUES (?,?,?,?,?,?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,Session("invoiceid")))
		objCmd.Parameters.Append(objCmd.CreateParameter("Name",200,1,30, your_name))
		objCmd.Parameters.Append(objCmd.CreateParameter("Rec_Name",200,1,30,rec_name))
		objCmd.Parameters.Append(objCmd.CreateParameter("Rec_Email",200,1,50, rec_email))
		objCmd.Parameters.Append(objCmd.CreateParameter("Message",200,1,250, message))
		objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,50,var_cert_code))
		objCmd.Parameters.Append(objCmd.CreateParameter("Amount",6,1,20, gift_amount))
		objCmd.Parameters.Append(objCmd.CreateParameter("Original_Amount",6,1,20, gift_amount))
		objCmd.Execute()

%>
	<!--#include virtual="emails/email_variables.asp"-->
<%
end if ' If there is a gift certificate in the order
%>