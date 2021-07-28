<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/functions/base64.asp"-->
<!--#include virtual="Connections/authnet.asp"-->
<!--#include virtual="/cart/inc_cart_main.asp" -->

<%
response.write "name: " & request.form("customer_name") & "<br>"
response.write "shipping_address1: " & request.form("shipping_address1") & "<br>"


		' Connect to auth.net
		strChargeCard = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
        & MerchantAuthentication() _
		& "<transactionRequest>" _
		    & "<transactionType>authCaptureTransaction</transactionType>" _
            & "<amount>12.00</amount>" _
            & "<payment>" _
                & "<opaqueData>" _
                    & "<dataDescriptor>COMMON.GOOGLE.INAPP.PAYMENT</dataDescriptor>" _
                    & "<dataValue>" & base64_encode(request.form("paymentToken")) & "</dataValue>" _
                & "</opaqueData>" _
            & "</payment>" _    
        & "</transactionRequest>" _
		& "</createTransactionRequest>"
		Set objResponseChargeCard = SendApiRequest(strChargeCard)

		RESPONSE.WRite "AUTH.NET RESPONSE:<br>" & objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text & "<br><br>"		
%>