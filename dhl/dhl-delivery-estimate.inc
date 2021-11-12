<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<%
Function getEstimatedDeliveryDate(shipping_type, address, city, state, zip, str_date_order_placed)
	
	date_order_placed = date()
	If str_date_order_placed <> "" Then date_order_placed = CDate(str_date_order_placed)
	
	If WeekDayName(WeekDay(date_order_placed)) = "Saturday" Then shipment_date = DateAdd("d", 2, date_order_placed)
	If WeekDayName(WeekDay(date_order_placed)) = "Sunday" Then shipment_date = DateAdd("d", 1, date_order_placed)
	If Time() > "08:00:00 AM" AND WeekDayName(WeekDay(date_order_placed)) <> "Saturday" AND WeekDayName(WeekDay(date_order_placed)) <> "Sunday" AND WeekDayName(WeekDay(date_order_placed)) <> "Friday" then
		shipment_date = DateAdd("d", 1, date_order_placed)
	End If
	If Time() > "08:00:00 AM" AND WeekDayName(WeekDay(date_order_placed)) = "Friday" then
		shipment_date = DateAdd("d", 3, date_order_placed)
	end if
	If Time() < "08:00:00 AM" AND WeekDayName(WeekDay(date_order_placed)) <> "Saturday" AND WeekDayName(WeekDay(date_order_placed)) <> "Sunday" then
		shipment_date = date_order_placed
	end if
	shipment_date = Year(shipment_date) & "-" &  Pd(Month(shipment_date),2) & "-" & Pd(DAY(shipment_date),2) 

	'====== IF REQUEST IS BEING MADE FROM /admin/dhl/dhl-request-label-v4.asp THEN REWRITE VARIABLE TO SHIP OUT TODAY. THE ONLY TIME LABELS ARE EVER REQUESTED IS THE DAY OF SHIPMENT =========
	if request_label_page = "yes" then 
		shipment_date = Year(date()) & "-" &  Pd(Month(date()),2) & "-" & Pd(DAY(date()),2)
	end if


	If address = "" Then address = Request("address")
	If city = "" Then city = Request("city")
	If state = "" Then state = Request("state")
	If zip = "" Then zip = Request("zip")

	json_dhl_shipments = """pickup"":""" & dhl_production_pickup_num &"""," & _
	"""distributionCenter"":""" & dhl_production_distribution_center & ""","  & _
	"""orderedProductId"":""" & shipping_type & """," & _	
	"""consigneeAddress"":{" & _
		"""name"":""Consignee""," & _
		"""address1"":""" & address & """," & _
		"""address2"":""""," & _
		"""city"":""" & city & """," & _
		"""state"":""" & state & """," & _
		"""postalCode"":""" & zip & """," & _
		"""country"":""US""," & _
		"""phone"":""""" & _
	"}," & _
	"""returnAddress"": {" & _
		"""name"": ""BAF""," & _
		"""address1"": ""1966 S Austin Ave""," & _
		"""city"": ""Georgetown""," & _
		"""state"": ""TX""," & _
		"""postalCode"": ""78626""," & _
		"""country"": ""US""" & _
	"},"  & _
	"""packageDetail"": {" & _
		"""packageId"": ""1""," & _
		"""packageDescription"": ""Body Jewelry""," & _
		"""weight"":{" & _
			"""value"": 8," & _
			"""unitOfMeasure"": ""OZ""" & _
			"}"  & _
		"},"  & _
	"""estimatedDeliveryDate"": {"  & _
			"""calculate"": true,"  & _
			"""expectedTransit"": 99,"  & _
			"""expectedShipDate"": """ & shipment_date & """"  & _
		"}"


	' =================== REQUEST DELIVERY TIMEFRAMES ====================================
	set rest = Server.CreateObject("Chilkat_9_5_0.Rest")
	'  Connect to the REST server.
	bTls = 1
	port = 443
	bAutoReconnect = 1
	success = rest.Connect(dhl_api_url,port,bTls,bAutoReconnect)
	success = rest.AddHeader("Content-Type","application/json")

	set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
	success = sbAuthHeaderVal.Append("Bearer ")
	success = sbAuthHeaderVal.Append(db_dhl_access_token)
	rest.Authorization = sbAuthHeaderVal.GetAsString()

	ResponseGetEstimates = rest.FullRequestString("POST","/shipping/v4/products", "{" & json_dhl_shipments & json_dhl_customs_info & "}")

	set JsonEstimates = Server.CreateObject("Chilkat_9_5_0.JsonObject")
	JsonEstimates.EmitCompact = 0
	JsonEstimates.Load(ResponseGetEstimates)
	'Response.Write "<pre>" & Server.HTMLEncode( JsonEstimates.Emit()) & "</pre>"

	if JsonEstimates.StringOf("products") <> "" then
		Set productsArray = JsonEstimates.ArrayOf("products")
		productsSize = productsArray.Size

		While i < productsSize           
			Set productsObj = productsArray.ObjectAt(i)
			mail_type = productsObj.StringOf("orderedProductId")
			If mail_type = "EXP" OR mail_type = "MAX" Then ' Domestic shipments
				If mail_type = "EXP" Then EXP_estimated_delivery = productsObj.StringOf("estimatedDeliveryDate.estimatedDeliveryMax")
				If mail_type = "MAX" Then MAX_estimated_delivery = productsObj.StringOf("estimatedDeliveryDate.estimatedDeliveryMax")		
			End IF
			i = i + 1
		Wend        
	End if
	If shipping_type = "EXP" Then getEstimatedDeliveryDate = EXP_estimated_delivery
	If shipping_type = "MAX" Then getEstimatedDeliveryDate = MAX_estimated_delivery
End Function
%>

<%
'=========== MAKE MONTH AND DAY IN DATE HAVE A LEADING 0 IF NEEDED ====================
Function pd(n, totalDigits) 
    if totalDigits > len(n) then 
        pd = String(totalDigits-len(n),"0") & n 
    else 
        pd = n 
    end if 
End Function 
'======================================================================================
%>