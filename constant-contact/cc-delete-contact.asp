<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/functions/base64.asp" -->
<!--#include virtual="/Connections/constant-contact.asp" -->
<!--#include virtual="/constant-contact/cc-validate-token.asp"-->
<%
' =========== USE THE REFRESH TOKEN TO GET A NEW ACCESS TOKEN ======================
if cc_validate_error = "unauthorized" then
%>
<!--#include virtual="/constant-contact/cc-refresh-access-token.asp"-->
<%
end if

'======== UNSUBSCRIBES CONTACT BASED ON EMAIL ADDRESS ===========================
Set objGetContact = Server.CreateObject("MSXML2.ServerXMLHTTP")
objGetContact.open "GET", "https://api.cc.email/v3/contacts?email=" & request.querystring("email"), false
objGetContact.SetRequestHeader "Authorization", "Bearer " & cc_access_token
objGetContact.setRequestHeader "Content-Type", "application/json"
objGetContact.Send()

jsonGetContactString  = objGetContact.responseText
Set oJSON = New aspJSON
oJSON.loadJSON(jsonGetContactString)

cc_contact_id = oJSON.data("contacts").item(0).item("contact_id") 
'response.write "<br/>Contact ID:" & cc_contact_id & "<br/>"
'response.write jsonGetContactString


'======== UNSUBSCRIBES CONTACT BASED ON CONTACT ID ===========================
Set objDeleteContact = Server.CreateObject("MSXML2.ServerXMLHTTP")
objDeleteContact.open "PUT", "https://api.cc.email/v3/contacts/" & cc_contact_id, false
objDeleteContact.setRequestHeader "Cache-Control", "no-cache"
objDeleteContact.SetRequestHeader "Authorization", "Bearer " & cc_access_token
objDeleteContact.setRequestHeader "Content-Type", "application/json"
objDeleteContact.setRequestHeader "Accept", "application/json"
objDeleteContact.Send("{" & _
    """email_address"": {" & _
    """address"":""" & request.querystring("email") & """," & _
    """permission_to_send"":""unsubscribed""," & _
    """opt_out_reason"":""No longer interested""" & _
    "}," & _
    """update_source"":""Contact""" & _
    "}")

jsonDeleteContactString  = objDeleteContact.responseText
Set oJSON = New aspJSON
oJSON.loadJSON(jsonDeleteContactString)

response.write jsonDeleteContactString

DataConn.Close()
Set DataConn = Nothing
%>