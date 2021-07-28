<%@LANGUAGE="VBSCRIPT" %>
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


if request.querystring("sandbox") = "ON" then
	session("sandbox") = "ON"
	session("sandbox_toggle") = "OFF"
elseif request.querystring("sandbox") = "OFF" then
	session("sandbox") = "OFF"
	session("sandbox_toggle") = "ON"
elseif request.querystring("sandbox") = "" and session("sandbox") = "ON" then
	session("sandbox") = "ON"
	session("sandbox_toggle") = "OFF"
elseif request.querystring("sandbox") = "" and session("sandbox") = "OFF" then
	session("sandbox") = "OFF"
	session("sandbox_toggle") = "ON"
else
	session("sandbox") = "OFF"
	session("sandbox_toggle") = "ON"
end if
%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!DOCTYPE HTML>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sandox admin</title>
</head>

<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
	Enable & disable sandbox testing
</h5>

Sandbox testing: <strong><%= session("sandbox") %></strong>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<a href="?sandbox=<%= session("sandbox_toggle") %>" id="sandbox_toggle">Turn <strong><%= session("sandbox_toggle") %></strong></a>
</div>
<div id="front_load"></div>
</body>
</html>
<script type="text/javascript">

	$(document).ready(function(){
	
		// Toggle sandbox front end load
		$('#sandbox_toggle').click(function(){
			$("#front_load").load("../sandbox.asp?sandbox=<%= session("sandbox_toggle") %>");
		});
	
	
	});
	</script>
<%
DataConn.Close()
%>