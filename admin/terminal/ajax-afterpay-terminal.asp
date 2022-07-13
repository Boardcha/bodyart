<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/Connections/afterpay-credentials.asp"-->
<%

Set objAfterPayTerminal = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
objAfterPayTerminal.open "POST", afterpay_url & "/payments/" & request.form("trans_id") & "/" & request.form("afterpay_payments"), false
objAfterPayTerminal.SetRequestHeader "Authorization", "Basic " & afterpay_api_credential & ""
objAfterPayTerminal.setRequestHeader "Accept", "application/json"
objAfterPayTerminal.setRequestHeader "Content-Type", "application/json"
objAfterPayCeckout.setRequestHeader "User-Agent", "Bodyartforms/1.0 (Custom Platform/1.0.0; ASP; Bodyartforms/" & afterpay_merchant_id & ") https://bodyartforms.com"

if request.form("afterpay_payments") = "refund" then
objAfterPayTerminal.Send("{" & _
            """amount"": {" & _
                """amount"":""" & FormatNumber(request.form("amount"), -1, -2, -2, -2) & """," & _
                """currency"":""USD""" & _
            "}" & _
        "}")
end if 

jsonAuthstring  = objAfterPayTerminal.responseText
Set oJSON = New aspJSON
oJSON.loadJSON(jsonAuthstring)

'response.write jsonAuthstring

if oJSON.data("refundId") <> "" then
%>
{  
    "status":"success",
    "reason":"Afterpay refund successful"
}
<%
end if


if oJSON.data("errorCode") <> "" then
%>
{  
    "status":"failed",
    "reason":"<%= oJSON.data("message") %>"
}
<%
end if
%>