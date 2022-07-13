<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/Connections/afterpay-credentials.asp"-->
<%
'=============  This endpoint creates a checkout that is used to initiate the afterpay payment process. Afterpay uses the information in the checkout request to assist with the consumer’s pre-approval process. ========================================================================

'====== Retrieve order information
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT *, ISNULL(customer_first,'') + ' ' + ISNULL(customer_last,'') as 'customer_name' FROM sent_items WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,20, Session("invoiceid")))
Set rsGetOrder = objCmd.Execute()

'====== Retrieve ordered items
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TBL_OrderSummary.ProductID, TBL_OrderSummary.DetailID, jewelry.picture, jewelry.picture_400, largepic, TBL_OrderSummary.qty, TBL_OrderSummary.item_price,  ISNULL(jewelry.title,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(TBL_OrderSummary.PreOrder_Desc,'') as 'item_name'  FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID WHERE InvoiceID = ? AND free = 0 ORDER BY OrderDetailID ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,20, Session("invoiceid")))

set rsGetOrderItems = Server.CreateObject("ADODB.Recordset")
rsGetOrderItems.CursorLocation = 3 'adUseClient
rsGetOrderItems.Open objCmd
total_items = rsGetOrderItems.RecordCount
loop_item = 1

While NOT rsGetOrderItems.EOF
    if loop_item < total_items then
        item_build_comma = ","
    else
        item_build_comma = ""
    end if

    
    items_build = items_build & _
        "{" & _
            """name"":""" & replace(rsGetOrderItems.Fields.Item("item_name").Value,"""", "") & """," & _
            """sku"":""" & rsGetOrderItems.Fields.Item("DetailID").Value & """," & _
            """quantity"": " & rsGetOrderItems.Fields.Item("qty").Value & "," & _
            """pageUrl"": ""https://bodyartforms.com/productdetails.asp?productid=" & rsGetOrderItems.Fields.Item("ProductID").Value & """," & _
            """imageUrl"": ""https://bodyartforms-products.bodyartforms.com/" & rsGetOrderItems.Fields.Item("largepic").Value & """," & _
            """price"": {" & _
                """amount"":""" & rsGetOrderItems.Fields.Item("item_price").Value & """," & _
                """currency"":""USD""" & _
            "}" & _
        "}" & item_build_comma

    loop_item = loop_item + 1
rsGetOrderItems.MoveNext()
Wend

var_domain = Request.ServerVariables("SERVER_NAME")

Set objAfterPayCeckout = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
objAfterPayCeckout.open "POST", afterpay_url & "/checkouts", false
objAfterPayCeckout.SetRequestHeader "Authorization", "Basic " & afterpay_api_credential & ""
objAfterPayCeckout.setRequestHeader "Accept", "application/json"
objAfterPayCeckout.setRequestHeader "Content-Type", "application/json"
objAfterPayCeckout.setRequestHeader "User-Agent", "Bodyartforms/1.0 (Custom Platform/1.0.0; ASP; Bodyartforms/" & afterpay_merchant_id & ") https://bodyartforms.com"
objAfterPayCeckout.Send("{" & _
        """amount"": {" & _
            """amount"":""" & FormatNumber(session("third_party_total"), -1, -2, -2, -2) & """," & _
            """currency"":""USD""" & _
        "}," & _
        """consumer"": {" & _
            """phoneNumber"":""" & rsGetOrder.Fields.Item("phone").Value & """," & _
            """givenNames"":""" & rsGetOrder.Fields.Item("customer_first").Value & """," & _
            """surname"":""" & rsGetOrder.Fields.Item("customer_last").Value & """," & _
            """email"":""" & rsGetOrder.Fields.Item("email").Value & """" & _
        "}," & _
        """shipping"": {" & _
            """name"":""" & rsGetOrder.Fields.Item("customer_name").Value & """," & _
            """line1"":""" & rsGetOrder.Fields.Item("address").Value & """," & _
            """area1"":""" & rsGetOrder.Fields.Item("city").Value & """," & _
            """region"":""" & rsGetOrder.Fields.Item("state").Value & """," & _
            """postcode"":""" & rsGetOrder.Fields.Item("zip").Value & """," & _
            """countryCode"":""US""" & _
        "}," & _
        """items"": [" & _
            items_build & _
        "]," & _
        """merchant"": {" & _
            """redirectConfirmUrl"":""https://" & var_domain & "/checkout-afterpay.asp?type=afterpay""," & _
            """redirectCancelUrl"":""https://" & var_domain & "/checkout-afterpay.asp?type=afterpay""" & _
        "}," & _
        """merchantReference"": """ & session("invoiceid") & """" & _
    "}")

jsonAuthstring  = objAfterPayCeckout.responseText
Set oJSON = New aspJSON
oJSON.loadJSON(jsonAuthstring)

'response.write jsonAuthstring

session("afterpay_checkout_token") =  oJSON.data("token")

'response.write "<br>TOKEN: " & session("afterpay_checkout_token")

%>
{
    "afterpay_token":"<%= oJSON.data("token") %>"
}
