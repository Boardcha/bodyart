<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Bodyartforms cart transfer"
    page_description = "Bodyartforms cart transfer"
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<%
'IF NO SESSION COOKIE FOUND OR THE COOKIE WAS FOUND BUT DIDN'T MATCH THEN
'RETRIEVE NEW SESSION ID AND SET IT IN THE COOKIE
'RE-WRITE GUEST USER INFORMATION IN THE DATABASE
%>
<!--#include virtual="/bootstrap-template/filters.asp" -->

add to google tag manager for page load
<BR/>
<br>
<a href="?session=619894133">Click here to load this page</a>

<br/>
<br/>
FIRST SCENARIO: A user comes and there is no cookie information. We can't take the risk that anyone can just type a session ID into the cookie and transfer a cart. There needs to be a way to verify that this user indeed owns the cart. We could send a hashed session id + the database salt via the querystring... and then if it matches allow the user to transfer the cart contents.
<br/>
<br/>
SECOND SCENARIO: There is a stored cookie session id, however it does not match the session id + salt that's in the database. There should be SOME entry in the database that will match even if the tab is showing a new cookie info. So as long as the database can find a match, it should be a legit cart and then the database entry can be replaced with the new session id + salt information.
<br/>
<br/>
THIRD SCENARIO: The hashed session + salt via the querystring matches what it's in the database and the user is allowed to proceed as usual.
<br/>
<br/>
FOURTH SCENARIO: Bots or users find the page and try to exploit it
<br/>
<br/>
<%
response.write "<br>Session ID " & Session.SessionID 
response.write "<br>var_guest_customer_id " & var_guest_customer_id
response.write "<br>request.cookies(cartSelector) " & request.cookies("cartSelector")
%>
<!--#include virtual="/bootstrap-template/footer.asp" -->
