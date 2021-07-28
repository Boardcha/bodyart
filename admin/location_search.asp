<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% response.Buffer = false %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

If request.form("location") <> "" and request.form("section") <> "" then
	var_operator = "AND"
else
	var_operator = "OR"
end if

if request.form("location") = "" then
	var_location = 1234564
else
	var_location = request.form("location")
end if
if request.form("section") = "" then
	var_section = "AZ"
else
	var_section = request.form("section") 
end if

	  set objCmd = Server.CreateObject("ADODB.Command")
	  objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
	  objCmd.CommandText = "SELECT jewelry.ProductID, picture, title, ProductDetail1, Gauge, Length, location, ID_Description FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID INNER JOIN TBL_Barcodes_SortOrder ON DetailCode = ID_Number WHERE location = ? " & var_operator & " ID_Description = ? ORDER BY ID_Description ASC, location ASC"
	  objCmd.Parameters.Append(objCmd.CreateParameter("location",3,1,12,var_location))
	  objCmd.Parameters.Append(objCmd.CreateParameter("section",200,1,20,var_section))
	  Set Recordset1 = objCmd.Execute()

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../CSS/admin.css" />
<title>Location item search</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="content-grey admin-content">
<table class="admin-table">
	<thead>
		<tr>
			<th></th>
			<th>Section</th>
			<th>Location</th>
			<th>Item</th>
		</tr>
	</thead>
<% 
If NOT Recordset1.EOF then
While NOT Recordset1.EOF 
%>
	<tr>
		<td>
			<a href="product-edit.asp?ProductID=<%= Recordset1.Fields.Item("ProductID").Value %>&info=less"><img src="http://bodyartforms-products.bodyartforms.com/<%=(Recordset1.Fields.Item("picture").Value)%>" width="90" height="90" /></a>
		</td>
		<td>
			<%=(Recordset1.Fields.Item("ID_Description").Value)%>
		</td>
		<td>
			<%=(Recordset1.Fields.Item("location").Value)%>
		</td>
		<td>	
			<%=(Recordset1.Fields.Item("title").Value)%>&nbsp;<%=(Recordset1.Fields.Item("Gauge").Value)%>&nbsp;<%=(Recordset1.Fields.Item("Length").Value)%>&nbsp;<%=(Recordset1.Fields.Item("ProductDetail1").Value)%>	
		</td>
	</tr>
<%
  Recordset1.MoveNext()
Wend
%>
</table>
<%

	Recordset1.Close()
	Set Recordset1 = Nothing
else
%>

No records found

<% 
End if %> 
</div>
</body>
</html>