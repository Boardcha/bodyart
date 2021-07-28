<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>

<%
' Show a $0 option for carts with just a gift certificate
if session("var_giftcert") <> "yes" and session("var_other_items") = 1 then

'ONLY RUN THIS SCRIPT IF THE COUNTRY IS USA =========================
If Request("Country") = "USA" then


'************************************************************************
' ECOMMERCEMAX SOLUTIONS http://www.ecommercemax.com
' Contact: info@ecommercemax.com
' Updated January 2008
'
' VERSION 2
'
' This package was autodelivered to you upon your order 
' using Ecommercemax's PAYPAL IPN MADE EASY SYSTEM
' Don't forget to check it out at http://www.ecommercemax.com/paypal_ipn_made_easy.asp
'


'*************************************************************************************
' IMPORTANT: YOU HAVE TO SPECIFY YOUR OWN ACCESS LICENSE NUMBER, USER ID AND PASSWORD 
' THAT YOU REGISTERED WITH UPS.
' To register an END-USER ACCOUNT with UPS, go to: 
'   https://www.ups.com/servlet/registration?loc=en_US_EC&returnto=http://www.ec.ups.com/ecommerce/techdocs/online_tools.html

' *** YOU MAY NOT RESELL THIS SCRIPT. ***

'*************************************************************************************

function ups_shipping_select( AccessLicenseNumber, UserID, Password, default_ups_code, pickuptypecode, customerclassification, shippercity, shipperstateprovincecode, shipperpostalcode, shippercountrycode, receivercity, receiverstateprovincecode, receiverpostalcode, receivercountrycode, residentialaddressindicator, packagingtypecode, shipmentweight, largepackageindicator, Pkg_Length, Pkg_Width, Pkg_Height )
Dim mydoc, xml_response, strXML, xmlhttp
	Dim str
	
	str = ""
	
	strXML="<?xml version='1.0'?><AccessRequest xml:lang='en-US'><AccessLicenseNumber>5C4E700FF5EE9540</AccessLicenseNumber>"
	strXML=strXML & "<UserId>bodyartforms</UserId><Password>forever333</Password></AccessRequest><?xml version='1.0'?>"
	strXML=strXML & "<RatingServiceSelectionRequest xml:lang='en-US'>"
	strXML=strXML & "<Request><TransactionReference><CustomerContext>Rating and Service</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>"
	strXML=strXML & "<RequestAction>Rate</RequestAction>"
	strXML=strXML & "<RequestOption>shop</RequestOption></Request>"
	strXML=strXML & "<PickupType><Code>01</Code></PickupType>"
	strXML=strXML & "<CustomerClassification><Code>01</Code></CustomerClassification>"
	strXML=strXML & "<Shipment>"
	strXML=strXML & "<Shipper>"
  	strXML=strXML & "<ShipperNumber>6F0869</ShipperNumber>" 
	strXML=strXML & "<Address>"
	strXML=strXML & "<city>Georgetown</city>" 
	strXML=strXML & "<StateProvinceCode>TX</StateProvinceCode>" 
	strXML=strXML & "<PostalCode>78626</PostalCode>"
	strXML=strXML & "<CountryCode>US</CountryCode>"	
	strXML=strXML & "</Address>"
	strXML=strXML & "</Shipper>"	
	strXML=strXML & "<ShipTo><Address>"
	strXML=strXML & "<city>" & receivercity & "</city>"
	strXML=strXML & "<StateProvinceCode>" & receiverstateprovincecode & "</StateProvinceCode>"
	strXML=strXML & "<PostalCode>" & receiverpostalcode & "</PostalCode>"	
	strXML=strXML & "<CountryCode>" & receivercountrycode & "</CountryCode>"	
  	strXML=strXML & "<ResidentialAddressIndicator>1</ResidentialAddressIndicator>"	
	strXML=strXML & "</Address></ShipTo>"
	strXML=strXML & "<Service><Code>" & "11" & "</Code></Service>"
	strXML=strXML & "<Package><PackagingType><Code>02</Code>"
	strXML=strXML & "<Description>Package</Description></PackagingType>"
	strXML=strXML & "<Description>Rate Shopping</Description>"	
	strXML=strXML & "<PackageWeight><Weight>1</Weight></PackageWeight>"
	
	If CDbl(Pkg_Length) > 0 Or CDbl(Pkg_Width) > 0 Or CDbl(Pkg_Height) > 0 Then
			strXML = strXML & "<Dimensions>"
			If CDbl(Pkg_Length) > 0 Then
					strXML = strXML & "<Length>7</Length>"
			End If
			If CDbl(Pkg_Width) > 0 Then
					strXML = strXML & "<Width>4</Width>"
			End If
			If CDbl(Pkg_Height) > 0 Then
					strXML = strXML & "<Height>10</Height>"
			End If
			strXML = strXML & "<Units>IN</Units>"
			strXML = strXML & "</Dimensions>"
	End If	
	
	if largepackageindicator <> "" then
  	strXML=strXML & "<LargePackageIndicator>1</LargePackageIndicator>"	
	end if
	strXML=strXML & "</Package>"	
	strXML=strXML & "<ShipmentServiceOptions>" & "" & "</ShipmentServiceOptions>"
	strXML=strXML & "<RateInformation><NegotiatedRatesIndicator /></RateInformation>"
	strXML=strXML & "</Shipment></RatingServiceSelectionRequest>" 
	
	Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP") 
	xmlhttp.Open "POST","https://onlinetools.ups.com/ups.app/xml/Rate?",false 
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
	xmlhttp.send strXML 
	
	xml_response = xmlhttp.responseText 
	
	Set mydoc=Server.CreateObject("Microsoft.xmlDOM") 
	mydoc.loadxml( xml_response ) 
	
'	response.Write(xml_response )
	'response.End()

	Set Response_NodeList = mydoc.documentElement.selectNodes("Response") 
	ups_result = Response_NodeList.Item(0).selectSingleNode("ResponseStatusCode").Text
	
	if ups_result <> 1 then 'NOT SUCCESSFULL?
		'ups_shipping_select = "UPS Result: " & Response_NodeList.Item(0).selectSingleNode("ResponseStatusDescription").Text
	'	ups_shipping_select = "UPS Result: " & Response_NodeList.Item(0).selectSingleNode("Error/ErrorDescription").Text
		exit function
	else
		Set RatedShipment_NodeList = mydoc.documentElement.selectNodes("RatedShipment") 
		total_node = RatedShipment_NodeList.length - 1 		
		x = 0		
		
' Create div table and header
%>
	
<%
ship_ups_loop = 50
var_address_str = UCASE(request("address"))

var_address_str = replace(var_address_str, "P.O.", "PO BOX")
var_address_str = replace(var_address_str, "P. O.", "PO BOX")

if InStr(var_address_str, "PO BOX") = 0 then

		do while x <= total_node
			ups_scode = RatedShipment_NodeList.Item(x).selectSingleNode("Service/Code").Text
			
			' Just in case a negotatiated rates node can't be returned ...
			Set node = RatedShipment_NodeList.Item(x).selectSingleNode("NegotiatedRates/NetSummaryCharges/GrandTotal/MonetaryValue")
				If node Is Nothing then
			
			  ups_amt   = RatedShipment_NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text + .50
			
			else
			
			  ' REMOVE NEGOTIATED RATES UNTIL WE FIGURE SOMETHING OUT ABOUT PRICE DISCREPANCIES  2/1/2011
			  'ups_amt   = RatedShipment_NodeList.Item(x).selectSingleNode("NegotiatedRates/NetSummaryCharges/GrandTotal/MonetaryValue").Text + .50
			  
			  ups_amt   = RatedShipment_NodeList.Item(x).selectSingleNode("TotalCharges/MonetaryValue").Text + .50

			end if ' can't find negotatiated rates node
			
			
			ups_gdays = RatedShipment_NodeList.Item(x).selectSingleNode("GuaranteedDaysToDelivery").Text				
			select case ups_scode
					case "01" 
					  ups_desc = "2) UPS Next Day Air" 
					  ups_type_import = "1DA"
					  ups_display_type = "Next Day Air"
					  ups_est_delivery = "1 business day"
					case "02" 
					  ups_desc = "2) UPS 2nd Day Air"
					  ups_type_import = "2DA"
					  ups_display_type = "2nd Day Air"
					  ups_est_delivery =  "2 business days"
					case "03" 
					  ups_type_import = "GND"
					  ups_desc = "2) UPS Ground"
					  ups_display_type = "Ground"
					  ups_est_delivery = "5-6 business days"
					case "07" 
					  ups_desc = "2) UPS Worldwide Express"
					  ups_type_import = "XPR"
					  ups_display_type = "Worldwide Express"
					  ups_est_delivery = "1-3 business days. Usually by Noon."
					case "08" 
					  ups_desc = "2) UPS Worldwide Expedited "
					  ups_type_import = "XPD"
					  ups_display_type = "Worldwide Expedited"
					  ups_est_delivery = "2-5 business days. Delivery by end of day."
					case "11" 
					  ups_desc = "2) UPS Canada Standard"
					  ups_type_import = "ST"
					  ups_display_type = "Canada standard"
					  ups_est_delivery = "Average 7-14 days"
					case "12" 
					  ups_desc = "2) UPS 3 Day Select"
					  ups_type_import = "3DS"
					  ups_display_type = "3 Day Select"
					  ups_est_delivery = "3 business days"
					case "13" 
					  ups_desc = "2) UPS Next Day Air Saver"
					  ups_type_import = "1DP"
					  ups_display_type =  "Next Day Air Saver"
					  ups_est_delivery =  "Next day - usually by 3:00 pm"
					case "14" 
					  ups_desc = "2) UPS Next Day Air Early AM"
					  ups_type_import = "1DM"
					  ups_display_type =  "Next Day Air Early AM"
					  ups_est_delivery =  "Next day - usually by 8:00 am"
					case "54" 
					  ups_desc = "2) UPS Worldwide Express Plus"
					  ups_type_import = "XDM"
					  ups_display_type = "Worldwide Express Plus"
					  ups_est_delivery = "1-3 business days. Usually by 9:00 am."
					case "59" 
					  ups_desc = "2) UPS 2nd Day Air AM"
					  ups_type_import = "2DM"
					  ups_display_type = "2nd Day Air AM"
					  ups_est_delivery = "2 days - usually by 10:30 am"
					case "65" 
					  ups_desc = "2) UPS Saver"
					  ups_type_import = "WXS"
					  ups_display_type =  "Worldwide Saver"
					  ups_est_delivery =  "1-3 days. Delivery by end of day."
			end select


		
			if isNumeric(ups_gdays) then
				if cint(ups_gdays) > 0 then
  				if ups_gdays > 1 then
					ups_AvgTime = " (Approx " & ups_gdays & " days to deliver after leaving BAF)"
					else
					ups_AvgTime = " (Approx " & ups_gdays & " day to deliver after leaving BAF)"
					end if
				end if
			end if
			if x = 0 then
				If Request("Country") = "USA" then
				varChecked = " checked" 
				selected = "selected" 
				end if
  			session("default_shipping_cost") = ups_amt     ' OPTIONAL: YOU CAN USE THESE LINES TO ACCESS THE RESULT OF THE UPS FUNCTION
				session("default_shipping_type") = ups_desc		' IF YOU NEED TO USE THE RESULTS ON THE SAME PAGE
			else
				If Request("Country") = "USA" then
				varChecked = ""  
				  selected = ""
				  end if
			end if	
			
			if ups_scode = default_ups_code then selected = "selected"
			
			if ups_type_import = "2DA" OR ups_type_import = "GND" OR ups_type_import = "XPR" OR ups_type_import = "XPD" OR ups_type_import = "3DS" OR ups_type_import = "1DP" OR ups_type_import = "WXS" then
			
'			if ups_type_import = "GND" then
			
'				str = str & "<input name='ShippingType' type='radio' value='" & ups_type_import & ",6.95," & ups_desc & "," & ups_amt - 0.50 & ",64'" & varChecked & ">  &nbsp;$6.95 - " & ups_display_type & "<br /><br />"
'			x = x + 1
			
'			else

ship_ups_loop = ship_ups_loop + 1
%>		
<label class="col-12 col-xs-12 col-sm-6 col-md-4 col-lg-6 col-xl-4 col-break1600-4 col-break1900-3 btn btn-light d-block btn-sm rounded-0 text-left <%= var_active %>" style="border: .75em solid #fff">
	<div class="btn-sm btn-outline-secondary border border-secondary text-center d-block my-1">Select this method</div>
	<input name="shipping-option" type="radio" value="<%= ups_type_import %>,<%= ups_amt %>,<%= ups_desc %>,<%= ups_amt - 0.50 %>,64, UPS" data-id="<%= ship_ups_loop %>" data-price="<%= ups_amt %>">
	<div class="float-left mr-2">
		<img class="my-1" src="/images/ups-logo.png" style="max-height:40px">
	</div>
	<div class="float-left d-block">
		<div class="d-block font-weight-bold">
			<%= formatcurrency(ups_amt) %> - <%= ups_display_type %>
		</div>
			<%= ups_est_delivery %>		
	</div>
</label>
<%
			x = x + 1
			
'			end if ' show pulled pricing if not ground	
			
			else
			
				str = str
				x = x + 1	
				
			end if ' filtering out certain types of shipping	
					
		loop
end if ' po box found do not display
%>
	
<%		
	end if 
				
	str = str & ""
	ups_shipping_select = str 
  
End function


end if 'ONLY RUN THIS SCRIPT IF THE COUNTRY IS USA =========================
%>
<% ' ---- only run if country is USA ---------------


If Request("Country") = "USA" then

	if request("country") <> "" then
		var_country = request("country")
	end if
	
	if request("city") <> "" then
		var_city = request("city")
	else
		var_city = ""
	end if
	
	if request("state") <> "" then
		varProvince = request("state")
	end if
	
	if request("zip") <> "" then
		var_zip = request("zip")
	end if
	

'************************************************************************
' ECOMMERCEMAX SOLUTIONS http://www.ecommercemax.com
' Contact: info@ecommercemax.com
' Updated January 2008
' VERSION 2


country_text = "AL:ALBANIA,DZ:ALGERIA*,AS:AMERICAN SAMOA,AD:ANDORRA,AO:ANGOLA,AI:ANGUILLA,AG:ANTIGUA &amp; BARBUDA,AR:ARGENTINA*,AM:ARMENIA*,AW:ARUBA,AU:AUSTRALIA*,AT:AUSTRIA*,AZ:AZERBAIJAN*,AP:AZORES*,BS:BAHAMAS,BH:BAHRAIN,BD:BANGLADESH*,BB:BARBADOS,BY:BELARUS*,BE:BELGIUM*,BZ:BELIZE,BJ:BENIN,BM:BERMUDA,BT:BHUTAN,BO:BOLIVIA,BL:BONAIRE,BA:BOSNIA*,BW:BOTSWANA,BR:BRAZIL*,VG:BRITISH VIRGIN ISLES,BN:BRUNEI,BG:BULGARIA*,BF:BURKINA FASO,BI:BURUNDI,KH:CAMBODIA,CM:CAMEROON,CA:CANADA*,IC:CANARY ISLANDS*,CV:CAPE VERDE,KY:CAYMAN ISLANDS,CF:CENTRAL AFRICAN REPUBLIC,TD:CHAD,CL:CHILE,CN:CHINA*,CO:COLOMBIA,CG:CONGO,CK:COOK ISLANDS,CR:COSTA RICA,HR:CROATIA*,CB:CURACAO,CY:CYPRUS*,CZ:CZECH REPUBLIC*,CD:DEMOCRATIC REPUBLIC OF CONGO,DK:DENMARK*,DJ:DJIBOUTI,DM:DOMINICA,DO:DOMINICAN REPUBLIC,EC:ECUADOR,EG:EGYPT,SV:EL SALVADOR,EN:ENGLAND*,GQ:EQUATORIAL GUINEA,ER:ERITREA,EE:ESTONIA*,ET:ETHIOPIA,FO:FAEROE ISLANDS*,FJ:FIJI,FI:FINLAND*,FR:FRANCE*,GF:FRENCH GUIANA,PF:FRENCH POLYNESIA,GA:GABON,GM:GAMBIA,GE:GEORGIA*,DE:GERMANY*,GH:GHANA,GI:GIBRALTAR,GR:GREECE*,GL:GREENLAND*,GD:GRENADA,GP:GUADELOUPE,GU:GUAM,GT:GUATEMALA,GG:GUERNSEY*,GN:GUINEA,GW:GUINEA-BISSAU,GY:GUYANA,HT:HAITI,HO:HOLLAND*,HN:HONDURAS,HK:HONG KONG,HU:HUNGARY*,IS:ICELAND*,IN:INDIA*,ID:INDONESIA*,IQ:IRAQ,IE:IRELAND,IL:ISRAEL*,IT:ITALY*,CI:IVORY COAST,JM:JAMAICA,JP:JAPAN*,JE:JERSEY*,JO:JORDAN,KZ:KAZAKHSTAN*,KE:KENYA,KI:KIRIBATI,KO:KOSRAE*,KW:KUWAIT,KG:KYRGYZSTAN*,LA:LAOS,LV:LATVIA*,LB:LEBANON,LS:LESOTHO,LR:LIBERIA,LI:LIECHTENSTEIN*,LT:LITHUANIA*,LU:LUXEMBOURG*,MO:MACAU,MK:MACEDONIA*,MG:MADAGASCAR,ME:MADEIRA*,MW:MALAWI,MY:MALAYSIA*,MV:MALDIVES,ML:MALI,MT:MALTA,MH:MARSHALL ISLANDS*,MQ:MARTINIQUE*,MR:MAURITANIA,MU:MAURITIUS,MX:MEXICO*,FM:MICRONESIA*,MD:MOLDOVA*,MC:MONACO*,MN:MONGOLIA*,MS:MONTSERRAT,MA:MOROCCO,MZ:MOZAMBIQUE,MP:N. MARIANA ISLANDS,NA:NAMIBIA,NP:NEPAL,NL:NETHERLANDS*,AN:NETHERLANDS ANTILLES,NC:NEW CALEDONIA,NZ:NEW ZEALAND*,NI:NICARAGUA,NE:NIGER,NG:NIGERIA,NF:NORFOLK ISLAND,NB:NORTHERN IRELAND*,NO:NORWAY*,OM:OMAN,PK:PAKISTAN*,PW:PALAU*,PA:PANAMA,PG:PAPUA NEW GUINEA,PY:PARAGUAY,PE:PERU,PH:PHILIPPINES*,PL:POLAND*,PO:PONAPE*,PT:PORTUGAL*,PR:PUERTO RICO*,QA:QATAR,RE:REUNION*,RO:ROMANIA*,RT:ROTA,RU:RUSSIA*,RW:RWANDA,SS:SABA,SP:SAIPAN,SM:SAN MARINO*,SA:SAUDI ARABIA*,SF:SCOTLAND*,SN:SENEGAL,CS:SERBIA AND MONTENEGRO*,SC:SEYCHELLES,SL:SIERRA LEONE,SG:SINGAPORE*,SK:SLOVAKIA*,SI:SLOVENIA*,SB:SOLOMON ISLANDS,ZA:SOUTH AFRICA*,KR:SOUTH KOREA*,ES:SPAIN*,LK:SRI LANKA*,NT:ST. BARTHELEMY,SW:ST. CHRISTOPHER,SX:ST. CROIX*,EU:ST. EUSTATIUS,UV:ST. JOHN *,KN:ST. KITTS &amp; NEVIS,LC:ST. LUCIA,MB:ST. MAARTEN,TB:ST. MARTIN,VL:ST. THOMAS*,VC:ST. VINCENT&#047;GRENADINES,SR:SURINAME,SZ:SWAZILAND,SE:SWEDEN*,CH:SWITZERLAND*,SY:SYRIA,TA:TAHITI,TW:TAIWAN*,TJ:TAJIKISTAN*,TZ:TANZANIA,TH:THAILAND*,TL:TIMOR LESTE,TI:TINIAN,TG:TOGO,TO:TONGA,VG:TORTOLA,TT:TRINIDAD &amp; TOBAGO,TU:TRUK*,TN:TUNISIA,TR:TURKEY*,TM:TURKMENISTAN*,TC:TURKS &amp; CAICOS ISLANDS,TV:TUVALU,UG:UGANDA,UA:UKRAINE*,UI:UNION ISLAND,AE:UNITED ARAB EMIRATES,GB:UNITED KINGDOM*,US:UNITED STATES*,UY:URUGUAY*,VI:US VIRGIN ISLANDS*,UZ:UZBEKISTAN*,VU:VANUATU,VA:VATICAN CITY STATE*,VE:VENEZUELA,VN:VIETNAM*,VR:VIRGIN GORDA,WL:WALES*,WF:WALLIS &amp; FUTUNA ISLANDS,WS:WESTERN SAMOA,YA:YAP*,YE:YEMEN,CS:YUGOSLAVIA*,ZM:ZAMBIA,ZW:ZIMBABWE"


country_arr = split(country_text,",")
country_num = ubound(country_arr)


  retrieve = request("retrieve") 
	AccessLicenseNumber = request("AccessLicenseNumber")
	UserID = request("UserID")
	Password = request("Password")
	
'	If var_country = "USA" OR var_country = "US" then
	varcountry = "US"
'	else
'	varcountry = RsGetUPSCountry.Fields.Item("Country_UPSCode").Value
'	end if

	receiverpostalcode = var_zip
	receivercity = UCASE(var_city)
	receiverstateprovincecode= UCASE(varProvince)
	receivercountrycode	= UCASE(varcountry)

	if trim(receivercountrycode) = "" then
		if NOT isnumeric(shipperpostalcode) or len(trim(shipperpostalcode))<5 then  
			push_error "Invalid Shipper PostalCode"
			end if
			receivercountrycode = "US"
	end if
	if trim(shippercountrycode) = "" then
		if NOT isnumeric(receiverpostalcode) or len(trim(receiverpostalcode))<5 then  
			push_error "Invalid Receiver's PostalCode"
			end if
   	shippercountrycode = "US"
	end if

	
	if error_str = "" then
 		shipping_options = ups_shipping_select( AccessLicenseNumber, UserID, Password, "03", pickuptypecode, customerclassification, shippercity, shipperstateprovincecode, shipperpostalcode, shippercountrycode, receivercity, receiverstateprovincecode, receiverpostalcode, receivercountrycode, residentialaddressindicator, packagingtypecode, shipmentweight, largepackageindicator, Pkg_Length, Pkg_Width, Pkg_Height )		
	end if

'	PUSH SHIPPING OPTIONS TO THE SCREEn
	response.write shipping_options
	
sub push_error(str)
	if error_str = "" then 
		  error_str = "&#8226; " & str
	else
		error_str = error_str & "<br>" & "&#8226; " & str
		response.write error_str
	end if
end sub

set RsGetUPSCountry = nothing



end if  ' ---- only run if country is USA ---------------

end if ' do not run if there's only a gift cert in cart
%>
