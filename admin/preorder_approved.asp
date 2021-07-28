<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

Dim rsGetPreorders
Dim rsGetPreorders_numRows

Set rsGetPreorders = Server.CreateObject("ADODB.Recordset")
rsGetPreorders.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPreorders.Source = "SELECT InvoiceID, ProductID, DetailID, qty, PreOrder_Desc, detail_code, title, ProductDetail1, OrderDetailID, Gauge, Length FROM dbo.QRY_OrderDetails WHERE customorder = 'yes' AND (shipped = 'PRE-ORDER APPROVED' or shipped = 'ON ORDER') AND item_ordered = 0 AND brandname = '" + Request.querystring("Company") + "' ORDER BY jewelry, InvoiceID ASC"
rsGetPreorders.CursorLocation = 3 'adUseClient
rsGetPreorders.LockType = 1 'Read-only records
rsGetPreorders.Open()

rsGetPreorders_numRows = 0

' Get pre-order companies
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT name FROM TBL_Companies WHERE preorder_status = 1"
Set rsGetCompanies = objCmd.Execute()
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetPreorders_numRows = rsGetPreorders_numRows + Repeat1__numRows
%>
<%
if request.form("FRMupdate") = "yes" then
temp = Replace( Request.Form("Checkbox"), "'", "''" ) 
varID = Split( temp, ", " ) 

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING

For i = 0 To UBound(varID) 

commUpdate.CommandText = "UPDATE dbo.QRY_OrderDetails SET shipped = 'ON ORDER', date_sent = '"& date() &"' WHERE OrderDetailID = " & varID(i) 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()

   Next

temp = Replace( Request.Form("Checkbox"), "'", "''" ) 
varID = Split( temp, ", " ) 

set commUpdate2 = Server.CreateObject("ADODB.Command")
commUpdate2.ActiveConnection = MM_bodyartforms_sql_STRING

For i = 0 To UBound(varID) 

commUpdate2.CommandText = "UPDATE dbo.QRY_OrderDetails SET item_ordered = 1, item_ordered_date = GETDATE() WHERE OrderDetailID = " & varID(i) 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate2.Execute()

   Next

   Response.Redirect("preorder_approved.asp?Company=Industrial Strength")
end if
%>
<html>
<head>

<title>Pre-order review &amp; approval</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body>
<!--#include file="admin_header.asp"-->
<SCRIPT LANGUAGE="JavaScript">
	function checkAll(field)
	{
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function uncheckAll(field)
	{
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
	}
	
	// URL change menu
	$(document).on("change", '#brandChange', function() { 
		console.log("test");
		location.href = $(this).val();
	});
</script>
<div class="p-3">
<h5>
	Pre-orders that need to be placed
</h5> 

<form action="" method="post" name="FRM_update" id="FRM_update">
	<div class="form-inline mb-2">
		<select class="form-control form-control-sm mr-3" name="brandChange" id="brandChange">
			<option value="#" selected>Select company</option>
		<% while not rsGetCompanies.eof %>
		<option value="preorder_approved.asp?Company=<%= rsGetCompanies.Fields.Item("name").Value %>"><%= rsGetCompanies.Fields.Item("name").Value %></option>
		<% rsGetCompanies.movenext()
		wend
		rsGetCompanies.movefirst()
		%>
		</select> 
			<%= Request.querystring("Company") %>

</div>

<input class="btn btn-sm btn-secondary searchbox" type="button" name="UnCheckAll" value="uncheck all" onClick="uncheckAll(document.FRM_update.Checkbox)">
<input class="btn btn-sm btn-secondary searchbox" type="button" name="CheckAll" value="check all" onClick="checkAll(document.FRM_update.Checkbox)">
<% If Not rsGetPreorders.EOF Or Not rsGetPreorders.BOF Then %>
			  <table class="table table-sm table-hover mt-4" style="border-collapse:collapse">
				  <thead class="thead-dark">
				<tr>
				  <th width="10%" style="background-color:#BDBDBD;border: 1px solid #000000">Invoice</th>
				  <th width="10%" style="background-color:#BDBDBD;border: 1px solid #000000" align="center">Qty</th>
				  <th width="10%" style="background-color:#BDBDBD;border: 1px solid #000000">Code</th>
				  <th width="70%" style="background-color:#BDBDBD;border: 1px solid #000000">Description</th>
				</tr>
			</thead>
				<% 
	While ((Repeat1__numRows <> 0) AND (NOT rsGetPreorders.EOF)) 
	%>
				  <tr>
					<td style="border: 1px solid #000000">
						<input name="Checkbox" type="checkbox" value="<%=(rsGetPreorders.Fields.Item("OrderDetailID").Value)%>">
						<a href="invoice.asp?ID=<%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%>"> <%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%></a></td>
					<td style="border: 1px solid #000000" align="center"><%=(rsGetPreorders.Fields.Item("qty").Value)%></td>
					<td style="border: 1px solid #000000"><%=(rsGetPreorders.Fields.Item("detail_code").Value)%>&nbsp;</td>
					<td style="border: 1px solid #000000"><%=Replace((rsGetPreorders.Fields.Item("title").Value), "PRE-ORDER ", "")%>&nbsp;<%=(rsGetPreorders.Fields.Item("gauge").Value)%>&nbsp;<%=(rsGetPreorders.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetPreorders.Fields.Item("ProductDetail1").Value)%><br>
					  Specs: <%=(rsGetPreorders.Fields.Item("PreOrder_Desc").Value)%></td>
				  </tr>
				  <% 
	  Repeat1__index=Repeat1__index+1
	  Repeat1__numRows=Repeat1__numRows-1
	  rsGetPreorders.MoveNext()
	Wend
	%>
	</table>
	<div class="text-center">
		<input class="btn btn-sm btn-purple" type="submit" name="Submit2" value="Set to ON ORDER">
		<input name="FRMupdate" type="hidden" id="FRMupdate" value="yes">
	</div>

<% End If ' end Not rsGetPreorders.EOF Or NOT rsGetPreorders.BOF %>
</form>
<% If rsGetPreorders.EOF And rsGetPreorders.BOF Then %>
	<div class="alert alert-danger">No pre-orders to review </div>
<% End If ' end rsGetPreorders.EOF And rsGetPreorders.BOF %>

</div>
</body>
</html>
<%
rsGetPreorders.Close()
Set rsGetPreorders = Nothing
%>
