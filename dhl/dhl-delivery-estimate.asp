<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<link href="/CSS/baf.min.css?v=120318" rel="stylesheet" type="text/css" />
<%
'=========== MAKE MONTH AND DAY IN DATE HAVE A LEADING 0 IF NEEDED ====================
Function pd(n, totalDigits) 
    if totalDigits > len(n) then 
        pd = String(totalDigits-len(n),"0") & n 
    else 
        pd = n 
    end if 
End Function 

json_dhl_shipments = """pickup"":""" & dhl_production_pickup_num &"""," & _
"""distributionCenter"":""" & dhl_production_distribution_center & ""","  & _
"""consigneeAddress"":{" & _
    """name"":""Amanda Bunch""," & _
    """address1"":""""," & _
    var_address2 & _
    """city"":""Round Rock""," & _
    """state"":""TX""," & _
    """postalCode"":""78665""," & _
    """country"":""US""," & _
    var_email & _
    """phone"":""8772235005""" & _
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
    """packageId"": ""GM123456""," & _
    """packageDescription"": ""Body Jewelry""," & _
    """weight"":{" & _
        """value"": 5," & _
        """unitOfMeasure"": ""OZ""" & _
        "},"  & _
    """service"": ""DELCON""," & _
    """shippingCost"":{" & _
        """currency"": ""USD""," & _
        """declaredValue"": 32.95," & _
        """dutiesPaid"": false" & _
        "}"  & _
    "},"  & _
"""rate"": {"  & _
        """calculate"": true,"  & _
        """rateDate"": """ & YEAR(Date()) & "-" & Pd(Month(date()),2) & "-" & Pd(DAY(date()),2)  & ""","  & _
        """currency"": ""USD"""  & _
    "},"  & _
"""estimatedDeliveryDate"": {"  & _
        """calculate"": true,"  & _
        """expectedTransit"": 10,"  & _
        """expectedShipDate"": ""2021-03-20"""  & _
    "}"

json_dhl_customs_info = ""
json_dhl_line_items = ""


'==== begin loop '
json_dhl_line_items = json_dhl_line_items & "{" & _
    """itemDescription"": ""test item ""," & _
    """countryOfOrigin"": ""US""," & _
    """hsCode"": ""7117.90.9000""," & _
    """packagedQuantity"": 1," & _
    """itemValue"": 1," & _
    """skuNumber"": ""test sku""," & _
    """currency"": ""USD""" & _
"},"
' ==== move next item in loop
' ==== repeat loop until done

if json_dhl_line_items <> "" then
    json_dhl_line_items = left(json_dhl_line_items,len(json_dhl_line_items)-1)
end if

json_dhl_customs_info = ",""customsDetails"": [" & _
    json_dhl_line_items & _
    "]"

    json_dhl_customs_info = "" '=========== resetting variable for testing

'response.write "JSON OUTPUT:<br>" & json_dhl_shipments & "<br><br>"
'response.write "LINE ITEMS:<br>" & json_dhl_line_items & "<br><br>"


' =================== REQUEST DELIVERY TIMEFRAMES ====================================
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

'  Connect to the REST server.
bTls = 1
port = 443
bAutoReconnect = 1
success = rest.Connect(dhl_api_url,port,bTls,bAutoReconnect)
success = rest.AddHeader("Content-Type","application/json")

' Set the Authorization property to "Bearer <token>"
	set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
	success = sbAuthHeaderVal.Append("Bearer ")
	success = sbAuthHeaderVal.Append(db_dhl_access_token)
	rest.Authorization = sbAuthHeaderVal.GetAsString()

ResponseGetEstimates = rest.FullRequestString("POST","/shipping/v4/products", "{" & _
json_dhl_shipments & json_dhl_customs_info & _
"}")

    set JsonEstimates = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    JsonEstimates.EmitCompact = 0
    JsonEstimates.Load(ResponseGetEstimates)
    'Response.Write "<pre>" & Server.HTMLEncode( JsonEstimates.Emit()) & "</pre>"


    if JsonEstimates.StringOf("products") <> "" then
    Set productsArray = JsonEstimates.ArrayOf("products")
    productsSize = productsArray.Size
%>
    [
<%
    i = 0
    Do While i < productsSize
            
    Set productsObj = productsArray.ObjectAt(i)
%>
    {
    "dhl_mail_type":"<%= productsObj.StringOf("orderedProductId") %>",  
    "dhl_max_expected_delivery":"<%= WeekDayName(WeekDay(productsObj.StringOf("estimatedDeliveryDate.estimatedDeliveryMax"))) %>, <%= MonthName(Month(productsObj.StringOf("estimatedDeliveryDate.estimatedDeliveryMax"))) %>&nbsp;<%= day(productsObj.StringOf("estimatedDeliveryDate.estimatedDeliveryMax")) %>",
    "dhl_max_delivery_days":"<%= productsObj.StringOf("estimatedDeliveryDate.deliveryDaysMax") %>",
    "dhl_mail_rate":"<%= productsObj.StringOf("rate.amount") %>"
    }
<%                       
    if i+1 < productsSize then
        response.write ","
    end if
    i = i + 1
    Loop        
%>
    ]
<%
    end if '==== IF PRODUCTS ARE FOUND IN JSON
%>





