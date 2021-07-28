<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/authnet.asp" -->
<%

'			& "<invoiceNumber>" & Session("invoiceid") & "</invoiceNumber>" _

		' Connect to auth.net
		strChargeCard = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "<transactionRequest>" _
		& "		<transactionType>authCaptureTransaction</transactionType>" _
		& "<amount>" & request.form("amount") & "</amount>" _
		& "<profile><customerProfileId>" & request.form("cim_account_id") & "</customerProfileId><paymentProfile><paymentProfileId>" & request.form("cim_billing_id") & "</paymentProfileId></paymentProfile></profile>" _
		& "<order>" _
		& "<description>Body jewelry</description>" _
		& "</order>" _
		& "</transactionRequest>" _
		& "</createTransactionRequest>"
		Set objResponseChargeCard = SendApiRequest(strChargeCard)
		
		' APPROVED - If REGISTERED customer order is APPROVED -----------------------------------
		If IsApiResponseSuccess(objResponseChargeCard) Then

		' 1 = Approved, 2 = Declined, 3 = Error, 4 = Held for Review
		if objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text = 1 then ' if approved
		
%>
		{  
			"status":"<div class='notice-eco'>CHARGE APPROVED</div>"
		}	
<% 
		else ' payment declined does not equal 1
						
			var_message = objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text	
%>
		{  
			"status":"<div class='notice-red'>DECLINED, <%= var_message %></div>"
		}
<% 			
			
		end if ' if payment is declined	-  does not equal 1
		
		else ' Also set a decline if an error is returned from auth.net
			var_message = "Problem with payment information"
%>
		{  
			"status":"<div class='notice-red'>DECLINED, <%= var_message %></div>"
		}
<% 		
			
		end if ' if an error response came back from Auth.net	


		
DataConn.Close()
%>