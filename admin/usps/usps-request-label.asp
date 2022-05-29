<%@LANGUAGE="VBSCRIPT" CodePage = 65001 %>
<!--#include virtual="/Connections/endicia.asp" -->
<%
'IIS should process this page as 65001 (UTF-8), responses should be 
'treated as 28591 (ISO-8859-1).
Response.CharSet = "ISO-8859-1"
Response.CodePage = 28591
%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/functions/random_integer.asp"-->
<!--#include virtual="/functions/iif.asp"-->
<!--#include virtual="/Connections/chilkat.asp" -->
<%
'******************** TESTING MODE FOR LABELS ********************************
testing = "NO"      'If YES, it creates fake labels. "NO" should be set for real labels.



'==================== REMOVE ALL BASE64 IMAGES FROM DATABASE THAT ARE OLD ============
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "UPDATE sent_items SET usps_base64_shipping_label = '' WHERE CAST(date_sent AS date) < CAST(GETDATE()-120 AS date)"
objCmd.Execute()


' =================== REQUEST SHIPPING LABEL =====================================  
if request.querystring("all") = "yes" OR request.querystring("single") ="" then
'==== GET ALL SHIPPING LABELS DURING BATCH PRINT
    sql_where = "(usps_package_id IS NULL OR usps_package_id = '') AND ship_code = N'paid' AND  shipped = N'Pending shipment' AND giftcert_flag = 0 AND (shipping_type LIKE '%USPS%')"
end if 
if request.querystring("single") = "yes" then
'==== REQUEST SINGLE LABEL TO PRINT
    sql_where = "ID = ?"
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING

objCmd.CommandText = "SELECT top 100 PERCENT " & _
    "ID AS OrderNumber, " & _
    "company AS ShipToCompany, " & _
    "ISNULL(customer_first, '') + ' ' + ISNULL(customer_last, '') AS ShipToName, " & _
    "REPLACE(address, '\', '/') AS ShipToAddress1, " & _
    "ISNULL(REPLACE(address2, '\', '/'), '') AS ShipToAddress2, " & _
    "city AS ShipToCity, " & _
    "CASE " & _
        "WHEN country = 'USA' THEN ISNULL(state, '') " & _
        "ELSE ISNULL(state, '') + ISNULL(province, '') " & _
    "END AS ShipToState, " & _
    "REPLACE(REPLACE(zip, '(', ''), ')', '') AS ShipToZip, " & _
    "CASE " & _
        "WHEN country = 'USA' THEN 'US' " & _
        "WHEN country = 'Australia' THEN 'AU' " & _
        "WHEN country = 'Austria' THEN 'AT' " & _
        "WHEN country = 'Belgium' THEN 'BE' " & _
        "WHEN country = 'Brazil' THEN 'BR' " & _
        "WHEN country = 'Canada' THEN 'CA' " & _
        "WHEN country = 'Denmark' THEN 'DK' " & _
        "WHEN country = 'England' THEN 'GB' " & _
        "WHEN country = 'Finland' THEN 'FI' " & _
        "WHEN country = 'France' THEN 'FR' " & _
        "WHEN country = 'Germany' THEN 'DE' " & _
        "WHEN country = 'Great Britain' THEN 'GB' " & _
        "WHEN country = 'Great Britain and Northern Ireland' THEN 'GB' " & _
        "WHEN country = 'Greece' THEN 'GR' " & _
        "WHEN country = 'Holland' THEN 'NL' " & _
        "WHEN country = 'Hong Kong' THEN 'HK' " & _
        "WHEN country = 'Hungary' THEN 'HU' " & _
        "WHEN country = 'Ireland' THEN 'IE' " & _
        "WHEN country = 'Israel' THEN 'IL' " & _
        "WHEN country = 'Italy' THEN 'IT' " & _
        "WHEN country = 'Japan' THEN 'JP' " & _
        "WHEN country = 'Latvia' THEN 'LV' " & _
        "WHEN country = 'Netherlands' THEN 'NL' " & _
        "WHEN country = 'New Zealand' THEN 'NZ' " & _
        "WHEN country = 'Norway' THEN 'NO' " & _
        "WHEN country = 'Portugal' THEN 'PT' " & _
        "WHEN country = 'Romania' THEN 'RO' " & _
        "WHEN country = 'Singapore' THEN 'SG' " & _
        "WHEN country = 'Slovakia' THEN 'SK' " & _
        "WHEN country = 'Korea' THEN 'KR' " & _
        "WHEN country = 'South Korea' THEN 'KR' " & _
        "WHEN country = 'Spain' THEN 'ES' " & _
        "WHEN country = 'Sweden' THEN 'SE' " & _
        "WHEN country = 'Switzerland' THEN 'CH' " & _
        "WHEN country = 'Thailand' THEN 'TH' " & _
        "WHEN country = 'United Kingdom' THEN 'GB' " & _
        "ELSE country " & _
    "END AS ShipToCountry, " & _
    "phone AS ShipToPhone, " & _
    "email AS ShipToEmail, " & _
    "shipping_type, " & _
    "CASE WHEN d.subtotal - (total_preferred_discount + total_gift_cert + total_coupon_discount + total_store_credit + total_free_credits) <= 0 THEN 1 ELSE d.subtotal - (total_preferred_discount + total_gift_cert + total_coupon_discount + total_store_credit + total_free_credits) END AS 'OrderValue', " & _
    "PackagedBy, " & _
    "autoclave, " & _
    "DiscountPercent, " & _
    "total_preferred_discount, " & _
    "total_coupon_discount, " & _
    "pay_method " & _
    "FROM sent_items AS O " & _
    "INNER JOIN " & _
    "(Select InvoiceID, SUM(qty * item_price) as subtotal " & _
    "FROM TBL_OrderSummary " & _
    "GROUP BY InvoiceID " & _
    ") as d ON O.ID = d.InvoiceID " & _
    " LEFT JOIN (SELECT DISTINCT DiscountCode, DiscountPercent FROM TBLDiscounts) AS C ON O.coupon_code = C.DiscountCode" & _
    " WHERE "  & sql_where
if request.querystring("single") = "yes" then
    '==== REQUEST SINGLE LABEL TO PRINT
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, request.querystring("invoiceid") ))
end if


Set rsGetOrder = objCmd.Execute()


'If an error is occured at one of labels, "Request Labels" button must be clicked again for unsuccesful labels
error_occured = false
var_success_message = "All USPS labels have been created succesfully!"

While NOT rsGetOrder.EOF
	military_shipment = ""
	MailClass = ""
	shipmentType= ""
	weight = ""
	MailpieceShape = ""

	var_packageid = rsGetOrder.Fields.Item("OrderNumber").Value

	if request.querystring("newlabel") = "yes" then
		var_packageid = var_packageid & "0000" & getInteger(4)
	end if

	'============================ Endicia API Shipment Types ============================
	'PriorityExpress
	'First
	'LibraryMail
	'MediaMail
	'ParcelSelect
	'RetailGround
	'Priority
	'PriorityMailExpressInternational
	'FirstClassMailInternational
	'FirstClassPackageInternationalService
	'PriorityMailInternational
	'USPSReturn

	if rsGetOrder.Fields.Item("shipping_type").Value = "USPS Priority mail heavy" then
		MailClass ="Priority"
		shipmentType="domestic"
		weight = 60
		MailpieceShape = "MediumFlatRateBox"
	elseif rsGetOrder.Fields.Item("shipping_type").Value = "USPS Priority mail" then	
		MailClass ="Priority"
		shipmentType="domestic"
		weight = 8	
		MailpieceShape = "FlatRateEnvelope"
	elseif rsGetOrder.Fields.Item("shipping_type").Value = "USPS Express mail" then	
		MailClass ="PriorityExpress"
		shipmentType="domestic"
		weight = 8
		MailpieceShape = "FlatRateEnvelope"
	elseif rsGetOrder.Fields.Item("shipping_type").Value = "USPS First Class Mail" then	
		MailClass ="First"
		shipmentType="domestic"
		weight = 8
		MailpieceShape = "Parcel"
	elseif rsGetOrder.Fields.Item("shipping_type").Value = "USPS Global priority mail" then	
		MailClass ="PriorityMailInternational"
		shipmentType="international"
		weight = 8
		MailpieceShape = "FlatRateEnvelope"
	elseif rsGetOrder.Fields.Item("shipping_type").Value = "USPS Express mail international" then	
		MailClass ="PriorityMailExpressInternational"
		shipmentType="international"
		weight = 8
		MailpieceShape = "FlatRateEnvelope"
	end if	

	'========= DETECT APO MILITARY ADDRESS ========
	if InStr(rsGetOrder.Fields.Item("ShipToState").Value, "AE") > 0 OR InStr(rsGetOrder.Fields.Item("ShipToCity").Value, "APO") > 0 OR InStr(rsGetOrder.Fields.Item("ShipToCity").Value, "FPO") > 0 OR InStr(rsGetOrder.Fields.Item("ShipToCity").Value, "DPO") > 0 then	
		military_shipment = "yes"
		'shipmentType="international"
	end if

	ShipToCompany = ""
	if Not IsNull(rsGetOrder.Fields.Item("ShipToCompany").Value) AND rsGetOrder.Fields.Item("ShipToCompany").Value<>"" then 
		ShipToCompany = "<ToCompany>" & rsGetOrder.Fields.Item("ShipToCompany").Value & "</ToCompany>"
	end if	
	
			
	' ENDICIA API Notes:
	' You can use the <PartnerCustomerID> to identify your customer’s requests using your account, 
	' and the <PartnerTransactionID> to uniquely identify your customer’s transaction, if you would like. Otherwise, you can simply set a default value for them, respectively 100, 200
	
		
	var_address2 = ""
	if rsGetOrder.Fields.Item("ShipToAddress2").Value <> "" then
		var_address2 = "<ToAddress2>" & rsGetOrder.Fields.Item("ShipToAddress2").Value & "</ToAddress2>"
	end if

	var_cleaned_zip = ""
	if rsGetOrder.Fields.Item("ShipToZip").Value <> "" then
		var_cleaned_zip = rsGetOrder.Fields.Item("ShipToZip").Value
		zip_str = Instr(var_cleaned_zip,"-")
		If zip_str Then var_cleaned_zip = Left(var_cleaned_zip, zip_str -1)
	end if

	XML="<?xml version=""1.0"" encoding=""utf-8""?>" & _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & _
					   "xmlns:xsd=""http://www.w3.org/2001/XMLSchema""" & _
					   "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
		  "<soap:Body>" & _
			"<GetPostageLabel xmlns=""www.envmgr.com/LabelService"">" & _
		 	"<LabelRequest " & IIF(shipmentType = "international","LabelType=""International"" LabelSubtype=""Integrated"" ImageRotation=""Rotate90""","") & " Test=""" & testing & """>" & _
				  "<MailClass>" & MailClass & "</MailClass>" & _
				  "<RubberStamp1>" & rsGetOrder("OrderNumber") & "</RubberStamp1>" & _
				  "<MailpieceShape>" & MailpieceShape & "</MailpieceShape>" & _
				  "<WeightOz>" & weight & "</WeightOz>" & _
				  "<RequesterID>" & endicia_requester_id & "</RequesterID>" & _
				  "<AccountID>" & endicia_account_id & "</AccountID>" & _
				  "<PassPhrase>" & endicia_passphrase & "</PassPhrase>" & _ 
				  "<PartnerCustomerID>100</PartnerCustomerID>" & _ 
				  "<PartnerTransactionID>200</PartnerTransactionID>" & _ 
				  ShipToCompany & _
				  "<ToName>" & rsGetOrder.Fields.Item("ShipToName").Value & "</ToName>" & _
				  "<ToAddress1>" & rsGetOrder.Fields.Item("ShipToAddress1").Value & "</ToAddress1>" & _
				  var_address2 & _
				  "<ToCity>" & rsGetOrder.Fields.Item("ShipToCity").Value & "</ToCity>" & _
				  "<ToState>" & rsGetOrder.Fields.Item("ShipToState").Value & "</ToState>" & _
				  "<ToPostalCode>" & var_cleaned_zip & "</ToPostalCode>" & _
				  "<FromCompany>Bodyartforms</FromCompany>" & _
				  "<FromName>BAF Bodyartforms</FromName>" & _
				  "<ReturnAddress1>1966 S Austin Ave</ReturnAddress1>" & _
				  "<FromCity>Georgetown</FromCity>" & _
				  "<FromState>TX</FromState>" & _
				  "<FromPostalCode>78626</FromPostalCode>" & _
				  "<FromEMail>service@bodyartforms.com</FromEMail>" & _
				  "<FromPhone>8772235005</FromPhone>"
	
	' IN INTERNATIONAL SHIPMENTS, API REQUIRES ENTERING DIMENSIONS OF THE PACKAGE
	' BELOW DIMENSIONS SHOULD BE MODIFIED WHEN STARTING TO SHIP INTERNATIONALLY
	If shipmentType="international" Then 
	' "<DateAdvance>0</DateAdvance>" & _
		XML = XML &	"<FromCountry>US</FromCountry>" & _
					"<ToCountryCode>" & rsGetOrder.Fields.Item("ShipToCountry").Value & "</ToCountryCode>" & _
					"<OriginCountry>CN</OriginCountry>" & _
					"<MailpieceDimensions>" & _
						"<Length>8</Length>" & _
						"<Width>6</Width>" & _
						"<Height>1</Height>" & _
					"</MailpieceDimensions>"
	end if ' internationals only

	If shipmentType="international" OR military_shipment = "yes" Then
		XML = XML &	"<CustomsCertify>TRUE</CustomsCertify>" & _
					"<CustomsSigner>Amanda Bunch</CustomsSigner>" & _
					"<CustomsInfo>" & _
						"<ContentsType>Merchandise</ContentsType>" & _
						"<CustomsItems>" 

		set objCmd = Server.CreateObject("ADODB.command")
		CursorLocation = 3
		objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
		objCmd.CommandText = "SELECT s.qty, s.item_price, d.ProductDetailID, LEFT(ISNULL(d.Gauge, '') + ' ' + ISNULL(d.Length, '') + ' ' + ISNULL(d.ProductDetail1, '') + ' ' + ISNULL(j.title, ''),50) AS description, SaleExempt, j.tariff_code, s.InvoiceID, j.ProductID FROM jewelry AS j INNER JOIN TBL_OrderSummary AS s ON j.ProductID = s.ProductID INNER JOIN ProductDetails AS d ON s.DetailID = d.ProductDetailID WHERE s.InvoiceID = ? AND s.item_price > 0"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, rsGetOrder.Fields.Item("OrderNumber").Value ))
		objCmd.ActiveConnection.CursorLocation = 3 'adUseClient
		Set rsGetLineItems = objCmd.Execute()

		While NOT rsGetLineItems.EOF
			itemExist = true
			'==== Set default value    
			calculated_item_price = rsGetLineItems.Fields.Item("item_price").Value

				if rsGetOrder.Fields.Item("DiscountPercent").Value > 0 AND rsGetLineItems.Fields.Item("SaleExempt").Value = 0 then
				
						calculated_item_price = ((100 - rsGetOrder.Fields.Item("DiscountPercent").Value) / 100) * rsGetLineItems.Fields.Item("item_price").Value
				
				end if '===== IF A DISCOUNT IS FOUND 
			

			' Total weights of the items must be equal or lighter than <WeightOz> tag
			item_weight = Round(weight / rsGetLineItems.RecordCount, 2) - 0.01
			XML = XML & _
				"<CustomsItem>" & _
					"<Description>" & replace(replace(TRIM(rsGetLineItems.Fields.Item("description").Value), """", ""), "Insert", "") & "</Description>" & _
					"<ContentsType>Merchandise</ContentsType>" & _
					"<Quantity>" & rsGetLineItems.Fields.Item("qty").Value & "</Quantity>" & _
					"<Weight>" & item_weight & "</Weight>" & _
					"<Value>" & FormatNumber(calculated_item_price, 2) & "</Value>" & _
				"</CustomsItem>"
				
			rsGetLineItems.MoveNext()
		Wend
		
		If itemExist <> true then
			'---- An item must be present to send out packageDescription
			
			XML = XML & _
				"<CustomsItem>" & _
					"<Description>REPLACEMENT OF LOST ITEM - Body Jewelry</Description>" & _
					"<Quantity>1</Quantity>" & _
					"<Weight>" & weight & "</Weight>" & _
					"<Value>1</Value>" & _
				"</CustomsItem>"			
		End if
		
		XML = XML & _
			"</CustomsItems>" & _
			  "</CustomsInfo>"
	end if 'shipmentType="international"


	XML = XML &		"</LabelRequest>" & _
				"</GetPostageLabel>" & _
			  "</soap:Body>" & _
			"</soap:Envelope>"
	
	'Debug
	'Response.Write XML
	
	set http = Server.CreateObject("Chilkat_9_5_0.Http")
	set objXml = Server.CreateObject("Chilkat_9_5_0.Xml")
	success = objXml.LoadXml(XML)
	If (success <> 1) Then
		errorOnGettingLabel = errorOnGettingLabel + 1
		xmlError = xmlError & "<br><pre>" & Server.HTMLEncode( objXml.LastErrorText) & "</pre>"
	End If

	strXml = objXml.GetXml()
	'response.write strXml

	' We'll need to add this in the HTTP header:
	' SOAPAction: "www.envmgr.com/LabelService/GetPostageLabel"
	http.SetRequestHeader "SOAPAction","www.envmgr.com/LabelService/GetPostageLabel"

	
	' Some services expect the content-type in the HTTP header to be "application/xml" while
	' other expect text/xml.  The default sent by Chilkat is "application/xml", 
	' but Endicia web service expects "text/xml".  Therefore, change the content-type:
	http.SetRequestHeader "Content-Type","text/xml; charset=utf-8"

	' The endpoint for this soap service is:
	endPoint = endicia_url

	' resp is a Chilkat_9_5_0.HttpResponse
	Set resp = http.PostXml(endPoint,strXml,"utf-8")
	If (http.LastMethodSuccess <> 1) Then
		httpError = httpError & "<br><pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre><br><br>"
	End If
	'response.write "<br><pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre><br><br>"
	
	set xmlResp = Server.CreateObject("Chilkat_9_5_0.Xml")
	success = xmlResp.LoadXML(resp.BodyStr)

	' Debug XML response
	'Response.Write "<pre>" & Server.HTMLEncode( xmlResp.GetXml()) & "</pre>"
	
	responseStatusCode = resp.StatusCode
	if responseStatusCode = 200 Then
	'response.write "TEST CODE "
		labelStatus = xmlResp.AccumulateTagContent("Status","")
		If labelStatus ="0" Then
			If shipmentType = "domestic" Then usps_base64_shipping_label = xmlResp.AccumulateTagContent("Base64LabelImage","") 
			If shipmentType = "international" Then usps_base64_shipping_label = xmlResp.AccumulateTagContent("Image","") 
			trackingNumber = xmlResp.AccumulateTagContent("TrackingNumber","") 
			amountPaid = xmlResp.AccumulateTagContent("FinalPostage","") 

			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
			objCmd.CommandText = "UPDATE sent_items SET USPS_tracking = ?, usps_package_id = ?, usps_amount_paid = ?, usps_base64_shipping_label = ? WHERE ID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("USPS_tracking", 200, 1, 100, trackingNumber))
			objCmd.Parameters.Append(objCmd.CreateParameter("usps_package_id", 200, 1, 100, trackingNumber))
			objCmd.Parameters.Append(objCmd.CreateParameter("usps_amount_paid",6, 1, 8, amountPaid))
			objCmd.Parameters.Append(objCmd.CreateParameter("usps_base64_shipping_label", 200, 1, -1, usps_base64_shipping_label))
			objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid", 3, 1, 12, rsGetOrder.Fields.Item("OrderNumber").Value))
			objCmd.Execute()
		Else
			error_occured = true
			If httpError<>"" Then var_request_error = var_request_error & httpError
			var_request_error = var_request_error & "Error on the invoice <a href='/admin/invoice.asp?ID=" & rsGetOrder.Fields.Item("OrderNumber").Value & "' target='_blank'>" & rsGetOrder.Fields.Item("OrderNumber").Value & "</a><br>Error Message: " & xmlResp.AccumulateTagContent("ErrorMessage","") & ""
		End If
	Else
		error_occured = true
		If httpError<>"" Then var_request_error = var_request_error & httpError
		var_request_error = var_request_error & "Error on the invoice <a href='/admin/invoice.asp?ID=" & rsGetOrder.Fields.Item("OrderNumber").Value & "' target='_blank'>" & rsGetOrder.Fields.Item("OrderNumber").Value & "</a>"
	End If
	
	rsGetOrder.MoveNext()
Wend

If error_occured = true Then 
	var_status = "error"
	var_message = var_request_error 
Else 
	var_status = "success"
	var_message = var_success_message
End If

%>
{
    "status":"<%= var_status %>",
    "message":"<%= var_message %>"
}