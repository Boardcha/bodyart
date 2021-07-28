<%@LANGUAGE="VBSCRIPT" CodePage = 65001 %>
<%
'IIS should process this page as 65001 (UTF-8), responses should be 
'treated as 28591 (ISO-8859-1).
Response.CharSet = "ISO-8859-1"
Response.CodePage = 28591
%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<!--#include virtual="/functions/random_integer.asp"-->

<%
'==================== REMOVE ALL BASE64 IMAGES FROM DATABASE THAT ARE OVER 1 WEEK OLD ============
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "UPDATE sent_items SET dhl_base64_shipping_label = '' WHERE CAST(date_sent AS date) < CAST(GETDATE()-120 AS date)"
objCmd.Execute()


' =================== REQUEST SHIPPING LABEL =====================================  
if request.querystring("all") = "yes" then
'==== GET ALL SHIPPING LABELS DURING BATCH PRINT
    sql_where = "(dhl_package_id IS NULL OR dhl_package_id = '') AND ship_code = N'paid' AND  shipped = N'Pending shipment' AND giftcert_flag = 0 AND (shipping_type LIKE '%DHL%')"
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
    " LEFT OUTER JOIN TBLDiscounts AS C ON O.coupon_code = C.DiscountCode" & _
    " WHERE "  & sql_where
if request.querystring("single") = "yes" then
    '==== REQUEST SINGLE LABEL TO PRINT
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, request.querystring("invoiceid") ))
end if

Set rsGetOrder = objCmd.Execute()

if rsGetOrder.EOF then
    var_status = "success"
    var_success_message = "No new orders to add. All labels have been created"
end if

While NOT rsGetOrder.EOF
var_packageid = rsGetOrder.Fields.Item("OrderNumber").Value

var_address2 = ""
if rsGetOrder.Fields.Item("ShipToAddress2").Value <> "" then
    var_address2 = """address2"":""" & rsGetOrder.Fields.Item("ShipToAddress2").Value & ""","
end if
if rsGetOrder.Fields.Item("shipping_type").Value = "DHL Expedited Max" then  
    if rsGetOrder.Fields.Item("ShipToState").Value = "AK" OR rsGetOrder.Fields.Item("ShipToState").Value = "HI" then  
        var_shipping_type = "EXP"
    else
        var_shipping_type = "MAX"
    end if
elseif rsGetOrder.Fields.Item("shipping_type").Value = "DHL Basic mail" then  
    var_shipping_type = "EXP"
elseif rsGetOrder.Fields.Item("shipping_type").Value = "DHL GlobalMail Packet Priority" OR rsGetOrder.Fields.Item("shipping_type").Value = "DHL Global basic ground" then  
    var_shipping_type = "PKY"
    var_packageid = "GM" & rsGetOrder.Fields.Item("OrderNumber").Value
elseif rsGetOrder.Fields.Item("shipping_type").Value = "DHL GlobalMail Parcel Priority" then  
    var_shipping_type = "PLY"
    var_packageid = "GM" & rsGetOrder.Fields.Item("OrderNumber").Value
else
    var_shipping_type = "EXP" ' Basic DHL mail
end if

if request.querystring("newlabel") = "yes" then
    var_packageid = var_packageid & "-" & getInteger(4)
end if

if rsGetOrder.Fields.Item("ShipToCountry").Value <> "USA" AND rsGetOrder.Fields.Item("ShipToCountry").Value <> "US" then 
    var_email = ""
else '===== domestic shipment
    var_email =  """email"":""" & rsGetOrder.Fields.Item("ShipToEmail").Value & ""","
end if

json_dhl_shipments = """orderedProductId"":""" & var_shipping_type & """," & _
    """consigneeAddress"":{" & _
    """name"":""" & rsGetOrder.Fields.Item("ShipToName").Value & """," & _
    """address1"":""" & rsGetOrder.Fields.Item("ShipToAddress1").Value & """," & _
    var_address2 & _
    """city"":""" & rsGetOrder.Fields.Item("ShipToCity").Value & """," & _
    """state"":""" & rsGetOrder.Fields.Item("ShipToState").Value & """," & _
    """postalCode"":""" & rsGetOrder.Fields.Item("ShipToZip").Value & """," & _
    """country"":""" & rsGetOrder.Fields.Item("ShipToCountry").Value & """," & _
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
    """packageId"": """ & var_packageid & """," & _
    """packageDescription"": ""Body Jewelry""," & _
    """weight"":{" & _
        """value"": 5," & _
        """unitOfMeasure"": ""OZ""" & _
    "},"  & _
    """serviceEndorsement"": ""4""," & _
    """shippingCost"":{" & _
        """currency"": ""USD""," & _
        """declaredValue"": " & rsGetOrder.Fields.Item("OrderValue").Value & "," & _
        """dutiesPaid"": false" & _
    "}"  & _
"}"

json_dhl_customs_info = ""
json_dhl_line_items = ""
if rsGetOrder.Fields.Item("ShipToCountry").Value <> "USA" AND rsGetOrder.Fields.Item("ShipToCountry").Value <> "US" then 

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    objCmd.CommandText = "SELECT s.qty, s.item_price, d.ProductDetailID, LEFT(ISNULL(d.Gauge, '') + ' ' + ISNULL(d.Length, '') + ' ' + ISNULL(d.ProductDetail1, '') + ' ' + ISNULL(j.title, ''),50) AS description, SaleExempt, j.tariff_code, s.InvoiceID, j.ProductID FROM jewelry AS j INNER JOIN TBL_OrderSummary AS s ON j.ProductID = s.ProductID INNER JOIN ProductDetails AS d ON s.DetailID = d.ProductDetailID WHERE s.InvoiceID = ? AND s.item_price > 0"
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, rsGetOrder.Fields.Item("OrderNumber").Value ))
    Set rsGetLineItems = objCmd.Execute()

    While NOT rsGetLineItems.EOF

        '==== Set default value    
        calculated_item_price = rsGetLineItems.Fields.Item("item_price").Value

            if rsGetOrder.Fields.Item("DiscountPercent").Value > 0 AND rsGetLineItems.Fields.Item("SaleExempt").Value = 0 then
            
                    calculated_item_price = ((100 - rsGetOrder.Fields.Item("DiscountPercent").Value) / 100) * rsGetLineItems.Fields.Item("item_price").Value
            
            end if '===== IF A DISCOUNT IS FOUND 

        json_dhl_line_items = json_dhl_line_items & "{" & _
            """itemDescription"": """ & replace(replace(TRIM(rsGetLineItems.Fields.Item("description").Value), """", ""), "Insert", "") & """," & _
            """countryOfOrigin"": ""US""," & _
            """hsCode"": """ & rsGetLineItems.Fields.Item("tariff_code").Value & """," & _
            """packagedQuantity"": " & rsGetLineItems.Fields.Item("qty").Value & "," & _
            """itemValue"": " & FormatNumber(calculated_item_price, 2) & "," & _
            """skuNumber"": """ & rsGetLineItems.Fields.Item("ProductDetailID").Value & """," & _
            """currency"": ""USD""" & _
        "},"
    rsGetLineItems.MoveNext()
    Wend
    
    if json_dhl_line_items <> "" then
        json_dhl_line_items = left(json_dhl_line_items,len(json_dhl_line_items)-1)
    else '---- An item must be present to send out packageDescription
        json_dhl_line_items = json_dhl_line_items & "{" & _
            """itemDescription"": ""REPLACEMENT OF LOST ITEM - Body Jewelry""," & _
            """countryOfOrigin"": ""US""," & _
            """hsCode"": ""7117.90.9000""," & _
            """packagedQuantity"": 1," & _
            """itemValue"": 1," & _
            """skuNumber"": ""Replacement""," & _
            """currency"": ""USD""" & _
        "}"    
    end if

    json_dhl_customs_info = ",""customsDetails"": [" & _
    json_dhl_line_items & _
    "]"

    'response.write "LINE ITEMS:----" & json_dhl_line_items & "----<br>"



end if

'RESPONSE.WRITE "FULL JSON STRING: <br>" & json_dhl_shipments & json_dhl_customs_info

'======== SENT REQUEST TO DHL FOR LABEL ===============================
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
    
    ResponseRequestLabel = rest.FullRequestString("POST","/shipping/v4/label?format=PNG", "{" & _
        """pickup"":""" & dhl_production_pickup_num &"""," & _
        """distributionCenter"":""" & dhl_production_distribution_center & """," & _
        json_dhl_shipments & json_dhl_customs_info & _
        "}")

    set JsonLabel = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    JsonLabel.EmitCompact = 0
    JsonLabel.Load(ResponseRequestLabel)
    'Response.Write "<pre>" & Server.HTMLEncode( JsonLabel.Emit()) & "</pre>"


if JsonLabel.StringOf("title") <> "" then '====== DISPLAY ERROR MESSAGE =============

    var_errors = ""

    IF JsonLabel.StringOf("title") = "Invalid Request" THEN

        var_request_error = "INVALID JSON REQUEST<br><br>{" & replace(json_dhl_shipments & json_dhl_customs_info, """", "'") & "}<br/><br>Invoice <a href='/admin/invoice.asp?ID=" & rsGetOrder.Fields.Item("OrderNumber").Value & "' target='_blank'>" & rsGetOrder.Fields.Item("OrderNumber").Value & "</a><br><br>" & var_request_error

    ELSEIF JsonLabel.StringOf("title") = "Access token expired" THEN

        var_request_error = "Access token expired" & var_request_error

    ELSE '=== If detailed errors are available =====

    

        Set errorsArray = JsonLabel.ArrayOf("invalidParams")
        errorsSize = errorsArray.Size

        For e = 0 To errorsSize - 1
            Set labelObj = errorsArray.ObjectAt(e)   

            var_errors = labelObj.StringOf("name") & " " &  labelObj.StringOf("reason") & "  |  " & var_errors

        Next

        var_request_error = var_errors & " - Invoice <a href='/admin/invoice.asp?ID=" & rsGetOrder.Fields.Item("OrderNumber").Value & "' target='_blank'>" & rsGetOrder.Fields.Item("OrderNumber").Value & "</a><br><br>" & var_request_error

    end if '==== detailed errors are available 
    
    var_status = "error"

else ' === IF SUCCESS   

    var_base64_image = ""
    
    var_dhl_package_id = JsonLabel.StringOf("labels[0].dhlPackageId")
    dhl_base64_shipping_label =  JsonLabel.StringOf("labels[0].labelData")

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    objCmd.CommandText = "UPDATE sent_items SET USPS_tracking = ?, dhl_package_id = ?, dhl_base64_shipping_label = ? WHERE ID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("tracking",200,1,500,var_dhl_package_id))
    objCmd.Parameters.Append(objCmd.CreateParameter("dhl_package_id",200,1,1000, var_dhl_package_id))
    objCmd.Parameters.Append(objCmd.CreateParameter("dhl_base64_shipping_label",200,1,-1, dhl_base64_shipping_label))
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12,rsGetOrder.Fields.Item("OrderNumber").Value))
    objCmd.Execute()

end if '====== ERROR NOT FOUND

rsGetOrder.MoveNext()
Wend
%>
{
    "status":"<%= var_status %>",
    "message":"<%= var_request_error & " " & var_success_message %>"
}
