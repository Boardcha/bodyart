<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes.asp" -->
<%
if request.form("auth") = "K3m!pl8r" and CustID_Cookie = 15 then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_Testimonials (Testimonial, Testimonial_Date, Testimonial_Name, Testimonial_Email) VALUES (?,'" & date() & "',?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("Testimonial",200,1,1500,request.form("testimonial")))
	objCmd.Parameters.Append(objCmd.CreateParameter("Testimonial_Name",200,1,30,request.form("name")))
	objCmd.Parameters.Append(objCmd.CreateParameter("Testimonial_Email",200,1,50,request.form("email")))
	objCmd.Execute()
	
	var_success = "yes"

end if
%>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO"
                        crossorigin="anonymous">
<meta charset="UTF-8">
<title>Add testimonial</title>
</head>

<body class="p-2">
<h5>Add a testimonial</h5>
	<% if var_success = "yes" then %>
	<div class="alert alert-success">Testimonial has been added!</div>
	<% end if %> 
	
<%
if request.querystring("auth") = "K3m!pl8r" then
%>
	<form name="form-add-testimonial" action="testimonial-add.asp" method="post">

	<div class="form-group">
		<label>Name</label>
		<input class="form-control" type="text" name="name" value="<%= request.querystring("name") %>" />
	</div>
	<div class="form-group">
		<label>Email</label>
		<input class="form-control" type="text" name="email" value="<%= request.querystring("email") %>" />
	</div>
	<div class="form-group">
		<label>Testimonial</label>
		<textarea class="form-control" rows="6" name="testimonial"><%= request.querystring("testimonial") %></textarea>
	</div>

	<input class="btn btn-purple btn-block" type="submit" name="button" value="Submit" />
	<input type="hidden" name="auth" value="<%= request.querystring("auth") %>">
	</form>
<% end if 'request.querystring("auth")
%>
</body>
</html>
