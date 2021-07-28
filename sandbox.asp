<%@LANGUAGE="VBSCRIPT" %>
<%
if request("sandbox") = "ON" then
	session("sandbox") = "ON"
	session("authnet_sandbox") = "ON"
	session("checkout_sandbox") = "ON"
	session("sandbox_toggle") = "OFF"
elseif request("sandbox") = "OFF" then
	session("sandbox") = "OFF"
	session("sandbox_toggle") = "ON"
	session("authnet_sandbox") = "OFF"
	session("checkout_sandbox") = "OFF"
	
elseif request("sandbox") = "" and session("sandbox") = "ON" then
	session("sandbox") = "ON"
	session("authnet_sandbox") = "ON"
	session("checkout_sandbox") = "ON"
	session("sandbox_toggle") = "OFF"
elseif request("sandbox") = "" and session("sandbox") = "OFF" then
	session("sandbox") = "OFF"
	session("authnet_sandbox") = "OFF"
	session("checkout_sandbox") = "OFF"
	session("sandbox_toggle") = "ON"
else
	session("sandbox") = "OFF"
	session("authnet_sandbox") = "OFF"
	session("checkout_sandbox") = "OFF"
	session("sandbox_toggle") = "ON"
end if
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Sandox toggle</title>
<style>
	body {
		font-size: 1em;
		font-family: verdana, sans-serif, arial;
		line-height: 150%;
	}
</style>
</head>

<body>
Database (front end): <strong>SANDBOX DATA</strong><br/>
Sandboxed interface (front end): <strong><%= session("sandbox") %></strong><br/>
Checkout sandbox (saving order): <strong><%= session("checkout_sandbox") %></strong><br/>
Auth.net sandbox: <strong><%= session("sandbox") %></strong><br/>
</body>
</html>
