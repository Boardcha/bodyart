<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../Connections/authnet.asp"-->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


' /checkout/inc-set-to-pending.asp is where orders are auto pushed through that are under $150 without comments etc
' the admin/ship_multiple_orders is where it pushes all of those through to ship out
' This page just displays the leftover orders that need to be manually reviewed

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM vw_sum_order_orderhistory WHERE shipped = 'Review' ORDER BY CASE WHEN city = 'Miami' THEN 1 WHEN country = 'Brazil' THEN 1 WHEN country = 'Hong Kong' THEN 1 WHEN giftcert_flag = 1 THEN 2 ELSE 3 END, ID ASC"

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.Open objCmd
rsGetRecords_total = rsGetRecords.RecordCount

' ===== Count all hidden records to be shipped
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM sent_items WHERE ship_code = 'paid' AND shipped = 'Pending...'"
set rsGetHiddenRecords = Server.CreateObject("ADODB.Recordset")
rsGetHiddenRecords.CursorLocation = 3 'adUseClient
rsGetHiddenRecords.Open objCmd
hidden_total = rsGetHiddenRecords.RecordCount
%>
<html>
<head>
<title>Review orders to be shipped</title>
<!--#include file="includes/inc_scripts.asp"-->
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

	// Approve hidden pre-approved orders
	$(document).on("click", "#approve-hidden-orders", function(event){
		$('#hidden-orders-message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');

		$.ajax({
		method: "post",
		dataType: "json",
		url: "invoices/ajax-approve-hidden-invoices.asp",
		data: $('#frm-hidden-orders').serialize()
		})
		.done(function(json,msg ) {
			if ($('#push_all').prop('checked')) {
				$.ajax({
					method: "post",
					url: "etsy/etsy-import-orders.asp"
					})
					.done(function(msg) {
						$("#hidden-orders-message").html('<span class="alert alert-success p-1">Items set to ship out</span>');
					})
					.fail(function(msg) {
						$("#hidden-orders-message").hide('<span class="alert alert-danger p-1">Failed importing etsy</span>');
					})
			} else {
				$("#hidden-orders-message").html('<span class="alert alert-success p-1">Items set to ship out</span>');
			}
			$('#hidden-order-count').html(json.records_total);
		})
		.fail(function(json,msg) {
			$("#hidden-orders-message").hide('<span class="alert alert-danger p-1">Failed importing regular orders</span>');
		});
	});
</script>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="mx-2">
<h5 class="mt-2"><%= rsGetRecords_total %> orders to be reviewed for shipment</h5>

<div class="container m-0 p-0" style="max-width:100%">
	<div class="row mb-2">
	  <div class="col text-left">
		  <div class="card bg-secondary text-light small">
			<div class="card-header p-2">
				<h6 class="p-0 m-0"><span id="hidden-order-count"><%= hidden_total %></span> pre-approved orders<span class="ml-3" id="hidden-orders-message"></span></h6>
			</div>
			<div class="card-body p-2">
				<form class="p-0 m-0" id="frm-hidden-orders">
					<div class="custom-control custom-radio  custom-control-inline">
						<input type="radio" id="push_all" name="push_hidden" class="custom-control-input" value="all" checked>
						<label class="custom-control-label" for="push_all">Ship ALL pre-approved orders</label>
					  </div>
					  <div class="custom-control custom-radio  custom-control-inline">
						<input type="radio" id="push_25" value="partial" name="push_hidden" class="custom-control-input">
						<label class="custom-control-label" for="push_25">Ship 25 pre-approved orders</label>
					  </div>
					  <button class="btn btn-sm btn-primary" type="button" id="approve-hidden-orders">Set to ship out</button>
				</form>
			</div>
		  </div>
	  </div>
	  <div class="col text-right">
		<form name="form" action="ship_multiple_orders.asp" method="post">
			<div>
				<input name="status" type="hidden" id="status" value="review"> 
				<input class="btn btn-sm btn-outline-secondary" type="button" name="UnCheckAll" value="Uncheck all" onClick="uncheckAll(document.form.Checkbox)">
				<input class="btn btn-sm btn-outline-secondary" type="button" name="CheckAll" value="Check all" onClick="checkAll(document.form.Checkbox)">
				<input class="btn btn-sm btn-primary ml-4" type="submit" name="Submit" value="Set to ship">
			</div>
	  </div>
	</div>
  </div>


<% If Not rsGetRecords.eof Then %>

<table class="table table-striped table-hover small">
<thead class="thead-dark">
            <tr>
              <th>Invoice #</th>
              <th>Name</th> 
              <th width="50%">Order description</th>
              <th></th>
              <th>Customer comments</th>
              <th width="3%"></th>
            </tr>
</thead class="thead-dark">
<tbody>
<% 
ii = 0
While NOT rsGetRecords.EOF

	add_class = ""
if instr(lcase(rsGetRecords.Fields.Item("city").Value), "miami") then
	'add_class = " table-danger "
end if
if instr(lcase(rsGetRecords.Fields.Item("country").Value), "brazil") OR instr(lcase(rsGetRecords.Fields.Item("country").Value), "hong kong") then
	'add_class = " table-danger "
end if
	class_giftcert = ""
if rsGetRecords.Fields.Item("giftcert_flag").Value = 1 then
	class_giftcert = " table-info "
end if

%>
	<tr class="<%= add_class %> <%= class_giftcert %>">
	  <td>
	  
		  <a href="invoice.asp?ID=<%= rsGetRecords.Fields.Item("ID").Value %>" target="_top" ><%=rsGetRecords.Fields.Item("ID").Value%></a>
		  <br/><br/>
		  <a href="order history.asp?var_first=<%=(rsGetRecords.Fields.Item("customer_first").Value)%>&var_last=<%=(rsGetRecords.Fields.Item("customer_last").Value)%>">View history</a>
		  <br/><br/>
		  <a href="email_template_send.asp?ID=<%=(rsGetRecords.Fields.Item("ID").Value)%>&type=generic">Email <%=(rsGetRecords.Fields.Item("customer_first").Value)%></a>
	  
	  </td>
	<td>
		
		<% if rsGetRecords.Fields.Item("company").Value <> "" then %><%=(rsGetRecords.Fields.Item("company").Value)%><br><% end if %>
		<%=(rsGetRecords.Fields.Item("customer_first").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("customer_last").Value)%><br>
		<%=(rsGetRecords.Fields.Item("address").Value)%><br>
				<% if (rsGetRecords.Fields.Item("address2").Value) <> "" then %><%=(rsGetRecords.Fields.Item("address2").Value)%><br><% end if %>
                <%=(rsGetRecords.Fields.Item("city").Value)%>, <%=(rsGetRecords.Fields.Item("state").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("province").Value)%> <%=(rsGetRecords.Fields.Item("zip").Value)%><br>
                <% if (rsGetRecords.Fields.Item("country").Value) = "USA" then %>
                <%=(rsGetRecords.Fields.Item("country").Value)%>
                <% else %>
                <font color="#000000" size="3" face="Century Gothic"><b><%=(rsGetRecords.Fields.Item("country").Value)%></b></font>
                <% end if %>
	</td> 
	
	<td><%=(rsGetRecords.Fields.Item("item_description").Value)%>
<%
Set rsGetOrderDetails = Server.CreateObject("ADODB.Recordset")
With rsGetOrderDetails
rsGetOrderDetails.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrderDetails.Source = "SELECT OrderDetailID, qty, title, ProductDetail1, item_price, notes, gauge, length  FROM dbo.QRY_OrderDetails  WHERE ID = " & rsGetRecords.Fields.Item("ID").Value & ""
rsGetOrderDetails.CursorLocation = 3 'adUseClient
rsGetOrderDetails.LockType = 1 'Read-only records
rsGetOrderDetails.Open()

Do While Not.Eof
%>
<%=(rsGetOrderDetails.Fields.Item("qty").Value)%> | <%=(rsGetOrderDetails.Fields.Item("title").Value)%>&nbsp; <%=(rsGetOrderDetails.Fields.Item("ProductDetail1").Value)%>&nbsp; <%=(rsGetOrderDetails.Fields.Item("Gauge").Value)%>&nbsp; <%=(rsGetOrderDetails.Fields.Item("Length").Value)%>&nbsp;&nbsp;&nbsp;<%= FormatCurrency(rsGetOrderDetails.Fields.Item("item_price").Value, -1, -2, -2, -2) %>&nbsp;<b><%=(rsGetOrderDetails.Fields.Item("notes").Value)%></b><br>
<%
	.Movenext()
Loop
End With 
%>
<br/>
<b><font size="3"><% if rsGetRecords.Fields.Item("total_less_discounts").Value < 0 then %>0<% else %><%= FormatCurrency(rsGetRecords.Fields.Item("total_less_discounts").Value, -1, -2, -0, -2) %><% end if %></font></b></td>
              <td>
<%=(rsGetRecords.Fields.Item("pay_method").Value)%><br>
<%=(rsGetRecords.Fields.Item("date_order_placed").Value)%><br>
<%=(rsGetRecords.Fields.Item("shipping_type").Value)%>&nbsp;&nbsp;&nbsp;<br>
<%=(rsGetRecords.Fields.Item("UPS_Service").Value)%><br>
Coupon: <%= rsGetRecords.Fields.Item("coupon_code").Value %>
<br>
<br>
<% if rsGetRecords.Fields.Item("pay_method").Value = "PayPal" then %>
Trans ID # <%= rsGetRecords.Fields.Item("transactionID").Value %>
<%
end if
%>
<% If rsGetRecords.Fields.Item("pay_method").Value <> "PayPal" then

' Authorize.net get transaction details
strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
& "<getTransactionDetailsRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
& MerchantAuthentication() _
& "<transId>" & rsGetRecords.Fields.Item("transactionID").Value & "</transId>" _
& "</getTransactionDetailsRequest>"

Set objGetTransactionDetails = SendApiRequest(strReq)

' If succcess retrieve transaction information
If IsApiResponseSuccess(objGetTransactionDetails) Then
	strAVSResponse = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:AVSResponse").Text


	If not(objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:cardCodeResponse") is nothing) then
			strCCVresponse = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:cardCodeResponse").Text
	End If


	If not(objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:payment/api:creditCard/api:cardNumber") is nothing) then
		strCardNumber = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:payment/api:creditCard/api:cardNumber").Text
	end if
'Else
'  Response.Write "The operation failed with the following errors:<br>" & vbCrLf
'  PrintErrors(objGetTransactionDetails)
End If


if strAVSResponse = "Y" or strAVSResponse = "X" then
	str_AVS_Friendly = "Street and zip both match"
elseif strAVSResponse = "A" then
	str_AVS_Friendly = "Only street matches, zip does not"
elseif strAVSResponse = "Z" or strAVSResponse = "W" then
	str_AVS_Friendly = "Only zip matches, street does not match"
elseif strAVSResponse = "N" then
	str_AVS_Friendly = "NO MATCH"
elseif strAVSResponse = "P" then
	str_AVS_Friendly = "AVS not applicable for this transaction"
elseif strAVSResponse = "U" then
	str_AVS_Friendly = "Address information is unavailable"
elseif strAVSResponse = "R" then
	str_AVS_Friendly = "Retry � System unavailable or timed out"
elseif strAVSResponse = "G" then
	str_AVS_Friendly = "Non-U.S. Card Issuing Bank"
elseif strAVSResponse = "B" then
	str_AVS_Friendly = "Address information not provided for AVS check"
elseif strAVSResponse = "S" then
	str_AVS_Friendly = "Service not supported by issuer"
else
	str_AVS_Friendly = "AVS Authorize.net system error"
end if

if strCCVresponse = "N"then
	str_CCV_Friendly = "NO MATCH"
elseif strCCVresponse = "M" then
	str_CCV_Friendly = "MATCH"
elseif strCCVresponse = "P" then
	str_CCV_Friendly = "Not processed"
elseif strCCVresponse = "S" then
	str_CCV_Friendly = "Should be on card, but is not indicated"
elseif strCCVresponse = "U" then
	str_CCV_Friendly = "Issuer is not certified or has not provided encryption key"
else
	str_CCV_Friendly = "Not processed"
end if

	if rsGetRecords.Fields.Item("customer_ID").Value <> 0 then
		registered_status = "Registered"
	else
		registered_status = "Not registered"
	end if
%>
<%= rsGetRecords.Fields.Item("IPaddress").Value %><br/>
<%= registered_status %><br>
<strong>AVS:</strong> <%= str_AVS_Friendly %><br/>
<strong>CVV:</strong> <%= str_CCV_Friendly %>

<% end if ' only show if paid with a credit card 

%>
                </td>
              <td class="text-primary"><%=(rsGetRecords.Fields.Item("customer_comments").Value)%></td>
              <td>
				<div class="custom-control custom-checkbox">
					<input type="checkbox" class="custom-control-input" id="customCheck<%= ii %>" name="Checkbox"  value="<%=(rsGetRecords.Fields.Item("ID").Value)%>">
					<label class="custom-control-label" for="customCheck<%= ii %>"></label>
				  </div>
              </td>
            </tr>
<% ii = ii + 1
rsGetRecords.MoveNext()
Wend
%>
</tbody>
          </table>
        
        <% End If ' end Not rsGetRecords.EOF Or NOT rsGetRecords.BOF %>
        
        <% If rsGetRecords.EOF And rsGetRecords.BOF Then %>
        <p class="faqs"><strong>No orders are available for review </strong></p>
        <% End If ' end rsGetRecords.EOF And rsGetRecords.BOF %>
        

<div class="text-right">
<input class="btn btn-sm btn-primary" type="submit" name="Submit" value="Set to ship">
</div>

</form>
</div>
</body>
</html>
<%
rsGetRecords.Close()
%>

