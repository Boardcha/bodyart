<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% response.Buffer = true %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if Request.QueryString("ProductID") <> "" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM jewelry WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,Request.QueryString("ProductID")))
	Set rs_getproduct = objCmd.Execute()
	
	noproduct = ""
	if rs_getproduct.BOF and rs_getproduct.EOF then
		page_title = "Product not found"
		noproduct = "Product not found"
	else
		page_title = rs_getproduct.Fields.Item("title").Value
	end if

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM ProductDetails INNER JOIN TBL_GaugeOrder ON COALESCE (ProductDetails.Gauge, '') = COALESCE (TBL_GaugeOrder.GaugeShow, '') INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ProductID = ? ORDER BY active DESC, item_order ASC, GaugeOrder ASC, price ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,Request.QueryString("ProductID")))
	Set rs_getdetails = objCmd.Execute()

end if 'if Request.QueryString("ProductID") <> ""	

%>
<!DOCTYPE html> 
<html>
<head>
<link rel="stylesheet" type="text/css" href="../CSS/print-friendly.css" />
<title>#<%= rs_getproduct.Fields.Item("ProductID").Value %>&nbsp;&nbsp;<%= page_title %></title>
<script type="text/javascript">
	window.print();
</script>
</head>
<body>
<img src="barcode.asp?code=<%= rs_getproduct.Fields.Item("ProductID").Value %>&height=30&width=1&mode=code39&text=0"><br/>
<img src="http://bodyartforms-products.bodyartforms.com/<%=(rs_getproduct.Fields.Item("picture").Value)%>">
<table style="border-spacing: 0">
<thead>
	<tr>
		<th>Detail #</th>
		<th>Section</th>
		<th>Location</th>
		<th>Qty</th>
		<th>Description</th>
	</tr>
</thead>
<% While NOT rs_getdetails.EOF  %>
<tbody>
	<tr>
		<td>
			<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>
		</td>
		<td>
			<%=(rs_getdetails.Fields.Item("ID_Description").Value)%>&nbsp;
			<% if rs_getdetails.Fields.Item("BinNumber_Detail").Value <> 0 then %>
			BIN <%=(rs_getdetails.Fields.Item("BinNumber_Detail").Value)%>
			<% end if %>
		</td>
		<td>
			<%= (rs_getdetails.Fields.Item("location").Value)%>
		</td>
		<td>
			<%=(rs_getdetails.Fields.Item("qty").Value)%>
		</td>
		<td>
			<% If (rs_getdetails.Fields.Item("Gauge").Value) <> "" Then %>
				<%= Server.HtmlEncode(rs_getdetails.Fields.Item("Gauge").Value)%>
			<% end if %>
			&nbsp;&nbsp;
			<% If (rs_getdetails.Fields.Item("Length").Value) <> "" Then %>
				<%= Server.HtmlEncode(rs_getdetails.Fields.Item("Length").Value)%>
			<% end if %>
			&nbsp;&nbsp;
			<% if rs_getdetails.fields.item("ProductDetail1").value <> "" then%>
				<%= Server.HTMLEncode(rs_getdetails.Fields.Item("ProductDetail1").Value)%>
			<% end if %>
		</td>
	</tr>
</tbody>
	<% 
	rs_getdetails.MoveNext()
	Wend
	%>
</table>
</body>
</html>
<%
Set rs_getuser = Nothing
Set rs_getdetails = Nothing
Set rs_getproduct = Nothing
DataConn.Close()
%>