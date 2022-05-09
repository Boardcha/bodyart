<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Unsubscribe"
	page_description = "Unsubscribe from Bodyartforms website emails"
	page_keywords = ""
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<!--#include virtual="/functions/encrypt.asp"-->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT customer_ID, salt FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
set rsGetUser = objCmd.Execute()

Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
password = "3uBRUbrat77V"
encrypt_link = rsGetUser("salt") & "___" & rsGetUser("customer_ID")

hashed_link_value = objCrypt.Encrypt(password, encrypt_link)
%>


<div class="display-5 mb-5">
	Unsubscribe
</div>
<a href="unsubscribe.asp?type=abandoned_cart&id=<%= hashed_link_value %>">Test delete link</a><br>
<% if request.querystring("type") = "abandoned_cart" then

'==== DECRYPT QUERYSTRING ID TO VERIFY BEFORE UPDATING CUSTOMER RECORD =============
    decrypt_link = request.querystring("id")
    decrypted = objCrypt.Decrypt(password, decrypt_link)

    decrypted_link_values = split(decrypted, "___")
    url_customer_salt = decrypted_link_values(0)
    url_customer_id = decrypted_link_values(1)

    '====== MATCH USER RECORD WITH DECRYPTED SALT AND CUSTOMER ID ================
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE customers SET unsubscribe_abandoned_cart_emails = 1 WHERE customer_ID = ? AND salt = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("customer_id",3,1,10, url_customer_id))
    objCmd.Parameters.Append(objCmd.CreateParameter("salt",200,1,100, url_customer_salt))
    set rsVerifyUser = objCmd.Execute()
%>
<div class="alert alert-success">You have been unsubscribed</div>
<% end if %>

<!--#include virtual="/bootstrap-template/footer.asp" -->
<%
Set objCrypt = Nothing
%>