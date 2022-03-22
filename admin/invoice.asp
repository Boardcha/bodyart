<%@LANGUAGE="VBSCRIPT" CodePage = 65001 %>
<% response.Buffer = true 
'IIS should process this page as 65001 (UTF-8), responses should be 
'treated as 28591 (ISO-8859-1).
Response.CharSet = "ISO-8859-1"
Response.CodePage = 28591
%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/authnet.asp"-->
<!--#include virtual="/functions/asp-json.asp"-->
<%'<!--#include virtual="/Connections/afterpay-credentials.asp"-->%>
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

if Request.Form("invoice_num") <> "" then
	var_invoiceid = Request.Form("invoice_num")
elseif Request.querystring("ID") <> "" then
	var_invoiceid = Request.querystring("ID")
else
	if request.querystring("create-empty-order") = "yes" then
		'==== CREATE EMPTY ORDER, Use the word "Empty" as a way to track the newest order ==========
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
		objCmd.CommandText = "INSERT INTO sent_items (shipped, date_order_placed, ship_code, pay_method, order_created_by) VALUES ('Pending...', '" & now() & "', 'paid', 'Empty', '" & user_name & "')"
		objCmd.Execute() 

		'===== RETRIEVE NEWEST EMPTY ORDER =================
		Set objCmd = Server.CreateObject ("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TOP(1) ID FROM sent_items WHERE pay_method = 'Empty' ORDER BY ID DESC" 
		Set rsGetEmptyOrder = objCmd.Execute()

		var_invoiceid = rsGetEmptyOrder("ID")

		'==== RESET EMPTY PAY METHOD TO BLANK ==========
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
		objCmd.CommandText = "UPDATE sent_items SET pay_method = '' WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,12, var_invoiceid))
		objCmd.Execute() 

	else
		var_invoiceid = 0
	end if
end if

if Request.Form("TransID") <> "" then
	var_transid = Request.Form("TransID")
	sql_trans = " OR transactionID = ?"
else
	var_transid = "123abc" ' fake id 
end if

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ? OR transactionID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,var_invoiceid))
objCmd.Parameters.Append(objCmd.CreateParameter("trans_id",200,1,50,var_transid))
Set rsGetOrder = objCmd.Execute()

if not rsGetOrder.eof then
	custID = rsGetOrder.Fields.Item("customer_ID").Value

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers WHERE customer_ID = '"&custID&"'"
Set rsGetCustomer = objCmd.Execute()

'if rsGetOrder.Fields.Item("coupon_code").Value <> "" then
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50,rsGetOrder.Fields.Item("coupon_code").Value))
	Set rsGetCouponDiscount = objCmd.Execute()
'end if

if rsGetOrder.Fields.Item("pay_method").Value <> "PayPal" then

	' Authorize.net get transaction details
	strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<getTransactionDetailsRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<transId>" & rsGetOrder.Fields.Item("transactionID").Value & "</transId>" _
	& "</getTransactionDetailsRequest>"
	
	Set objGetTransactionDetails = SendApiRequest(strReq)

	' If succcess retrieve transaction information
	If IsApiResponseSuccess(objGetTransactionDetails) Then
		strAVSResponse = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:AVSResponse").Text

		If not(objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:payment/api:creditCard/api:cardNumber") is nothing) then
			strCardNumber = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:payment/api:creditCard/api:cardNumber").Text
		end if
		
		' CVV status comes right before the invoice # and right after the AVS response. It's a one letter response in the api:transaction text
				
		If not(objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:cardCodeResponse") is nothing) then
			 strCCVresponse = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:cardCodeResponse").Text
		End If
	
		
	Else ' if there's an error getting a transaction
	'  Response.Write "The operation failed with the following errors:<br>" & vbCrLf
	 ' PrintErrors(objGetTransactionDetails)
	End If
	
else ' get paypal transaction details

	' PayPal/Authnet GET DETAILS OF TRANSACTION
	strPayPalGetDetails = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "  <transactionRequest>" _
	& "		<transactionType>getDetailsTransaction</transactionType>" _
	& "  	<refTransId>" & rsGetOrder.Fields.Item("transactionID").Value & "</refTransId>" _
	& "</transactionRequest>" _
	& "</createTransactionRequest>"
	Set objResponseDetails = SendApiRequest(strPayPalGetDetails)
	
	
	

	' PayPal request made
	If IsApiResponseSuccess(objResponseDetails) Then
	
'	Response.Write "Raw response: " & Server.HTMLEncode(objResponseDetails.selectSingleNode("/*/api:transactionResponse").text) & "<br><br>" & vbCrLf
		
		If not(objResponseDetails.selectSingleNode("/*/api:transactionResponse/api:secureAcceptance/api:PayerEmail") is nothing) then
			var_paypal_email = objResponseDetails.selectSingleNode("/*/api:transactionResponse/api:secureAcceptance/api:PayerEmail").Text
		end if
		If not(objResponseDetails.selectSingleNode("/*/api:transactionResponse/api:secureAcceptance/api:PayerID") is nothing) then
			var_paypal_id = objResponseDetails.selectSingleNode("/*/api:transactionResponse/api:secureAcceptance/api:PayerID").Text
		end if
		
	else ' show error for paypal transaction
					
		var_message = "<div class=""notice-red"">" & objResponseDetails.selectSingleNode("/*/api:transactionResponse").Text & "</div>"
		
	end if ' GET DETAILS OF PAYPAL TRANSACTION

end if ' if order is paid with credit card  or paypal

Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TBL_OrderSummary.InvoiceID, TBL_OrderSummary.ProductID, TBL_OrderSummary.DetailID, jewelry.picture, TBL_OrderSummary.qty, TBL_OrderSummary.PreOrder_Desc, jewelry.title, ProductDetails.ProductDetail1, ProductDetails.qty AS stock_qty, ProductDetails.price, sent_items.ID, TBL_OrderSummary.OrderDetailID, jewelry.customorder, sent_items.shipped, jewelry.brandname, TBL_OrderSummary.backorder, TBL_OrderSummary.item_shipped, TBL_OrderSummary.item_ordered, TBL_OrderSummary.item_received, TBL_OrderSummary.item_price, TBL_OrderSummary.notes, TBL_OrderSummary.discount, sent_items.email, anodization_id_ordered, sent_items.customer_ID, sent_items.ship_code, sent_items.price AS OrderTotal, TBL_OrderSummary.status, sent_items.Comments_OrderError, sent_items.Review_OrderError, ProductDetails.wlsl_price, sent_items.date_sent, TBL_OrderSummary.CharityPaid, TBL_OrderSummary.CharityPaidDate, TBL_OrderSummary.ProductReviewed, jewelry.jewelry, TBL_OrderSummary.anodization_fee, jewelry.SaleExempt, jewelry.type, ProductDetails.ProductDetailID,  ProductDetails.location, ProductDetails.DetailCode, sent_items.customer_first, TBL_OrderSummary.ProductPhotographed, TBL_Barcodes_SortOrder.ID_SortOrder, TBL_Barcodes_SortOrder.ID_Description, TBL_Barcodes_SortOrder.ID_Number, jewelry.type, ProductDetails.BinNumber_Detail, ProductDetails.Gauge, ProductDetails.Length, jewelry.active AS ActiveMain, ProductDetails.active AS ActiveDetail, TBL_OrderSummary.BackorderReview, ProductDetails.free, TBL_OrderSummary.item_problem, TBL_OrderSummary.TimesScanned, TBL_OrderSummary.ErrorReportDate, TBL_OrderSummary.ErrorDescription, TBL_OrderSummary.ErrorOnReview,  TBL_OrderSummary.ErrorSealedBag, TBL_OrderSummary.ErrorConserveBags, TBL_OrderSummary.ErrorQtyMissing, sent_items.PackagedBy, addon_item, TBL_OrderSummary.date_added, ISNULL(replace(jewelry.type,'None',''),'') + ' ' + ISNULL(jewelry.title,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') as 'item_description', TBL_OrderSummary.returned, TBL_OrderSummary.returned_qty, preorder_timeframes, item_ordered_date, item_received_date  FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number LEFT OUTER JOIN TBL_Companies ON jewelry.brandname = TBL_Companies.name WHERE InvoiceID = '" & rsGetOrder.Fields.Item("ID").Value & "' ORDER BY OrderDetailID ASC"
Set rsGetOrderItems = objCmd.Execute()


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

end if 'if not rsGetOrder.eof then
%>

<html>
<head>
<% if not rsGetOrder.eof then

if rsGetOrder.Fields.Item("customer_last").Value <> "" then
	customer_last = Server.HTMLEncode(rsGetOrder.Fields.Item("customer_last").Value)
end if
 %>
	<title><%=(rsGetOrder.Fields.Item("customer_first").Value)%>&nbsp;<%= customer_last %>&nbsp;-&nbsp;<%=(rsGetOrder.Fields.Item("ID").Value)%></title>
<% else %>
<title>No order found</title>
<% end if %>
<link rel="stylesheet" type="text/css" href="../CSS/redactor.css" />
<script type="text/javascript" src="../js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
<script type="text/javascript" src="../js/clipboard.js"></script>
<script type="text/javascript" src="scripts/invoice_auto_update_fields.js?v=012020"></script>
</head>
<body>
<!--#include file="admin_header.asp"-->
<% if not rsGetOrder.eof then%>
<input type="hidden" name="main-id" id="main-id" value="<%=(rsGetOrder.Fields.Item("ID").Value)%>" />
<%End If
if request("notice-type") = "deduct" OR request("notice-type") = "add" then

	notice_show = ""
	
	if request("notice-type") = "deduct" then
		notice_text = "Deducted inventory successfully"
	end if
	if request("notice-type") = "add" then
		notice_text = "Added inventory back into stock"
	end if
%>
<script type="text/javascript">
	var invoiceid = $('#main-id').val();
	$(window).on('load',function(){
		$('#modal-inventory-alert').modal('show');
		// Stops the alert modal from showing more than once by rewriting the URL state after page load.
		window.history.replaceState({}, document.title, "/admin/" + "invoice.asp?id=" + invoiceid);
	});
</script>
<%
end if

If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE 
if not rsGetOrder.eof then 
%>
<div class="modal fade" id="modal-inventory-alert" tabindex="-1" role="dialog">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title">Inventory status</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body">
				<div class="alert alert-success">
						<%= notice_text %>
				</div>
			</div>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		  </div>
		</div>
	  </div>


<div class="container mt-3 ajax-update" style="max-width:100%">
		<div class="row">
	<div class="col-sm pr-4 small">       
			<div class="container w-100">
				<div class="row">
					<div class="col-12 p-0 h4"><%=(rsGetOrder.Fields.Item("ID").Value)%>
						<a class="btn btn-sm btn-outline-secondary d-inline-block" href="invoice.asp?ID=<%= rsGetOrder.Fields.Item("ID").Value - 1%>"><i class="fa fa-angle-left fa-lg"></i></a>
						<a class="btn btn-sm btn-outline-secondary d-inline-block"  href="invoice.asp?ID=<%= rsGetOrder.Fields.Item("ID").Value + 1%>"><i class="fa fa-angle-right fa-lg"></i></a> 
					</div>
					<div class="col12 p-0 text-right">
						<a class="btn btn-sm btn-secondary d-inline-block" href="invoices/print-friendly-invoice.asp?ID=<%=(rsGetOrder.Fields.Item("ID").Value)%>" target="_blank">Print invoice</a>
						<span id="holder-copy-order"></span>
						<div id="label-message"></div>
					</div>
				</div>
			</div>     			
		
			<div class="form-group">
				<label class="font-weight-bold" for="status">Order status:</label>     
					<select name="status" id="order-status" data-column="shipped" data-friendly="Order status" class="ajax_input_fadeout form-control form-control-sm">
						<option value="<%=(rsGetOrder.Fields.Item("shipped").Value)%>" selected ><%=(rsGetOrder.Fields.Item("shipped").Value)%></option>
						<option value="Cancelled">Cancelled</option>
						<option value="CHARGEBACK">Chargeback</option>
						<option value="FLAGGED">Flagged order</option>
						<option value="Review">On review (pending shipment)</option>
						<option value="ON HOLD">Order on hold</option>
						<option value="Waiting for PayPal eCheck to clear">Paypal eCheck</option>
						<option value="PACKAGE CAME BACK">Package came back</option>
						<option value="Pending...">Pending...</option>
						<option value="Lost package">Lost package</option>
						<option value="CUSTOM ORDER IN REVIEW">Custom items in review</option>
						<option value="CUSTOM ORDER APPROVED">Custom items approved to order</option>
						<option value="ON ORDER">Custom items have been ordered</option>
						<option value="CUSTOM COLOR IN PROGRESS">Items need anodizing</option>
						<option value="RETURN">Return (Waiting for return)</option>
						<option value="RETURN (EXCEPTION)">Return (New orders allowed)</option>
						<option value="Shipped">Shipped</option>
					</select>
				</div>
				<div class="form-group">
				<label for="paid" class="font-weight-bold">Paid:</label>
				<select name="paid" data-column="ship_code" data-friendly="Order paid" class="ajax_input_fadeout form-control form-control-sm">
					<option value="<%=(rsGetOrder.Fields.Item("ship_code").Value)%>" selected ><%=(rsGetOrder.Fields.Item("ship_code").Value)%></option>
					<option value="paid">Paid</option>
					<option value="not paid">Not paid</option>
				</select>
			</div>
			<div class="form-group">
				<label class="font-weight-bold" for="date_shipped">Date shipped:</label>
				<span class="ml-5">Placed on <%=(rsGetOrder.Fields.Item("date_order_placed").Value)%></span>
				<input class="form-control form-control-sm" name="date_shipped" type="text" data-column="date_sent" data-friendly="Date sent" value="<%=(rsGetOrder.Fields.Item("date_sent").Value)%>">
			</div>
			<style>
					.icon_package {background-repeat:no-repeat; background-position: right; padding-right: 5px;
						background-image: url(../images/icon_package.png)}
			</style>
			<div class="form-group">
				<label class="font-weight-bold" for="packagedby">Assigned to packer:</label>
				<span class="ml-5">Pulled by <%=(rsGetOrder.Fields.Item("pulled_by").Value)%></span>
					<input class="form-control form-control-sm <% if (rsGetOrder.Fields.Item("ScanInvoice_Timestamp").Value) <> "" then %>icon_package<% end if %>" name="packagedby" type="text"  data-column="PackagedBy" data-friendly="Packaged by" value="<% = (rsGetOrder.Fields.Item("PackagedBy").Value) %>">
			</div>
			<div class="form-group">
					<label class="font-weight-bold" for="private_notes">Private notes:</label>
					<textarea class="form-control form-control-sm" rows="4" name="private_notes" id="private_notes"  data-column="our_notes" data-friendly="Private notes"  maxlength="250" placeholder="Add a new private note"></textarea><br/>
					<div id="display_notes">
					</div>
			</div>

 
	</div>
	<div class="col-sm px-4 border-right border-left small">  
		<div class="container w-100">
			<div class="row">
				<div class="col-8 p-0 h4">Shipping information</div>
				<div class="col-4 p-0 text-right">
						<a href="order history.asp?var_first=<%=(rsGetOrder.Fields.Item("customer_first").Value)%>&var_last=<%= customer_last %>" target="_blank" class="btn btn-sm btn-secondary">Order history</a>
				</div>
			</div>
		</div>      
		<div class="d-block my-2">
		<% if (rsGetOrder.Fields.Item("company").Value) <> "" then %>
		<%=(rsGetOrder.Fields.Item("company").Value)%><br>
		<% end if %>
		<%=(rsGetOrder.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsGetOrder.Fields.Item("customer_last").Value)%><br>
		<%=(rsGetOrder.Fields.Item("address").Value)%> <br>
		<% if (rsGetOrder.Fields.Item("address2").Value) <> "" then %>
		<%=(rsGetOrder.Fields.Item("address2").Value)%> <br>
		<% end if %>
		<%=(rsGetOrder.Fields.Item("city").Value)%>, <%=(rsGetOrder.Fields.Item("state").Value)%><%=(rsGetOrder.Fields.Item("province").Value)%>&nbsp;&nbsp;<%=(rsGetOrder.Fields.Item("zip").Value)%><br>
		<%=(rsGetOrder.Fields.Item("country").Value)%>
	</div>
		<span id="copy_address" class="clipboard btn btn-sm btn-secondary" data-clipboard-text="<% if (rsGetOrder.Fields.Item("company").Value) <> "" then %><%=(rsGetOrder.Fields.Item("company").Value)%>&#10;<% end if %><%=(rsGetOrder.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsGetOrder.Fields.Item("customer_last").Value)%>&#10;<%=(rsGetOrder.Fields.Item("address").Value)%><% if (rsGetOrder.Fields.Item("address2").Value) <> "" then %>&nbsp;<%=(rsGetOrder.Fields.Item("address2").Value)%><% end if %>&#10;<%=(rsGetOrder.Fields.Item("city").Value)%>, <%=(rsGetOrder.Fields.Item("state").Value)%><%=(rsGetOrder.Fields.Item("province").Value)%>&nbsp;&nbsp;<%=(rsGetOrder.Fields.Item("zip").Value)%>&#10;<%=(rsGetOrder.Fields.Item("country").Value)%>">Copy address</span>
		<span id="show_address" class="btn btn-sm btn-secondary">Edit address</span>
		<span id="hide_address" class="btn btn-sm btn-secondary" style="display:none">Close address info</span>
		<div id="div_shipping_address" class="alert alert-secondary mt-2" style="display:none">
				<div class="form-group">
						<label for="customerID"><span id="temp-account" class="pointer" data-custid="<%= rsGetOrder.Fields.Item("customer_ID").Value %>">*</span> Customer ID #</label>
						<input class="form-control form-control-sm"  name="customerID" type="text" data-column="customer_ID" data-friendly="Customer ID" value="<%=(rsGetOrder.Fields.Item("customer_ID").Value)%>">
							</div>
					<% if rsGetOrder.Fields.Item("customer_ID").Value <> 0 then 
					if NOT rsGetCustomer.EOF then %>
			<div class="alert alert-info">
					<a class="btn btn-sm btn-info" href="customer_edit.asp?ID=<%=(rsGetCustomer.Fields.Item("Customer_ID").Value)%>" target="_blank">Edit</a><br>
						<%=(rsGetCustomer.Fields.Item("customer_first").Value)%>&nbsp; <%=(rsGetCustomer.Fields.Item("customer_last").Value)%><br>
						<%=(rsGetCustomer.Fields.Item("email").Value)%> <br>
			<div class="form-group">
			  <label for="customer_credits">Credits $</label>
			  <input class="form-control form-control-sm"  name="customer_credit" id="customer_credit" type="text" data-column="credits" data-friendly="Customer credit"  value="<%=(rsGetCustomer.Fields.Item("credits").Value)%>" data-custid="<%= rsGetCustomer.Fields.Item("Customer_ID").Value %>">
						</div> 
			  <div id="confirm_credit_update" class="notice-eco d-none">Credit updated</div>
				</div>
					<% end if  
					end if %>
				<div class="form-group">
			<label for="email">E-mail</label>
			<input class="form-control form-control-sm" name="email" type="text" data-column="email" data-friendly="E-mail" value="<%=(rsGetOrder.Fields.Item("email").Value)%>">	
				</div>
				<div class="form-group">	
			<label for="company">Company</label>
			<input class="form-control form-control-sm"  name="company" type="text" data-column="company" data-friendly="Company" value="<%=(rsGetOrder.Fields.Item("company").Value)%>">	
				</div>
				<div class="form-group">
			<label for="first_name">First</label>
			<input class="form-control form-control-sm"  name="first_name" type="text" data-column="customer_first" data-friendly="First name" value="<%=(rsGetOrder.Fields.Item("customer_first").Value)%>">	
				</div>
				<div class="form-group">
			<label for="last_name">Last</label>
			<input class="form-control form-control-sm"  name="last_name" type="text" data-column="customer_last" data-friendly="Last name" value="<%=(rsGetOrder.Fields.Item("customer_last").Value)%>">
				</div>
		<div class="form-group">
			<label for="address">Address</label>
			<input class="form-control form-control-sm" name="address" type="text" data-column="address" data-friendly="Address" value="<%=(rsGetOrder.Fields.Item("address").Value)%>">
		</div><div class="form-group">
			<label for="address2">Address</label>
			<input class="form-control form-control-sm" name="address2" type="text" data-column="address2" data-friendly="Address" value="<%=(rsGetOrder.Fields.Item("address2").Value)%>">
		<//div>
		<div class="form-group">
			<label for="city">City</label>
			<input class="form-control form-control-sm" name="city" type="text" data-column="city" data-friendly="City" value="<%=(rsGetOrder.Fields.Item("city").Value)%>">
		</div>
		<div class="form-group">
			<label for="state">State</label>
			<input class="form-control form-control-sm" name="state" type="text" data-column="state" data-friendly="State" value="<%=(rsGetOrder.Fields.Item("state").Value)%>">
		</div>
		<div class="form-group">
			<label for="province">Province</label>
			<input class="form-control form-control-sm" name="province" type="text" data-column="province" data-friendly="Province" value="<%=(rsGetOrder.Fields.Item("province").Value)%>">
		</div>
		<div class="form-group">
			<label for="zip">Zip</label>
			<input class="form-control form-control-sm" name="zip" type="text" data-column="zip" data-friendly="Zip" value="<%=(rsGetOrder.Fields.Item("zip").Value)%>">
		</div>
		<div class="form-group">
			<label for="country">Country</label>
			<input class="form-control form-control-sm" name="country" type="text" data-column="country" data-friendly="Country" value="<%=(rsGetOrder.Fields.Item("country").Value)%>">
		</div>
		<div class="form-group">
			<label for="phone">Phone</label>
			<input class="form-control form-control-sm" name="phone" type="text" data-column="phone" data-friendly="Phone" value="<%=(rsGetOrder.Fields.Item("phone").Value)%>">
		</div>
			
		</div><!-- alert -->
	</div>
		
		<div class="form-group mt-3">

				<label class="font-weight-bold" for="shipping-type">Shipping type:</label>
				<select class="form-control form-control-sm" name="shipping-type" id="shipping-type" data-column="shipping_type" data-friendly="Shipping type" class="ajax_input_fadeout">
					<option value="<%=(rsGetOrder.Fields.Item("shipping_type").Value)%>" selected><%=(rsGetOrder.Fields.Item("shipping_type").Value)%></option>
					<option value="DHL Basic mail">DHL Basic mail (Domestic)</option>
					<option value="DHL Expedited Max">DHL Expedited Max (Domestic)</option>
					<option value="DHL GlobalMail Packet Priority">DHL GlobalMail Packet Priority (Tracked to border)</option>
					<option value="DHL GlobalMail Parcel Priority">DHL GlobalMail Parcel Priority (FULL TRACKING)</option>
					<option value="USPS First Class Mail">USPS First Class Mail</option>
					<option value="USPS Priority mail">USPS Priority mail</option>
					<option value="USPS Priority mail heavy">USPS Priority mail heavy</option>
					<option value="USPS Express mail">USPS Express mail</option>
					<option value="USPS Express mail international">USPS Express mail international</option>
					<option value="UPS ground">UPS ground - GND</option>
					<option value="UPS next day">UPS next day</option>
					<option value="UPS 2 day">UPS 2 day</option>
					<option value="UPS 3 day">UPS 3 day</option>
					<option value="UPS worldwide express">UPS worldwide express</option>
					<option value="OFFICE PICK UP">OFFICE PICK UP</option>
				</select>
			</div>
			<div class="form-group">
			<label class="font-weight-bold" for="shipping-cost">Shipping price paid:</label>
			<input class="form-control form-control-sm" name="shipping-cost" type="text" data-column="shipping_rate" data-friendly="Shipping cost" value="<%= (rsGetOrder.Fields.Item("shipping_rate").Value)%>">
			</div>
			<div class="form-group">
			<label for="ups_code" class="ups-tracking font-weight-bold">UPS Code</label>
			<input name="ups_code" id="ups_code" class="ups-tracking form-control form-control-sm" type="text" data-column="UPS_Service" data-friendly="UPS Code" value="<%=(rsGetOrder.Fields.Item("UPS_Service").Value)%>">
			</div>
			<div class="form-group">
			<label for="usps-tracking" class="usps-tracking font-weight-bold">USPS Tracking:</label>
			<input name="usps-tracking" id="usps-tracking" class="usps-tracking form-control form-control-sm" data-column="USPS_tracking" data-friendly="USPS tracking" type="text" value="<%=(rsGetOrder.Fields.Item("USPS_tracking").Value)%>">
			</div>
			<div class="mb-3">
				<% if rsGetOrder.Fields.Item("dhl_base64_shipping_label").Value = "" AND ISNULL(rsGetOrder.Fields.Item("dhl_base64_shipping_label").Value)  then 
				display_shipping_label = "display:none!important"
			end if
			if instr(rsGetOrder.Fields.Item("shipping_type").Value,"DHL") > 0 then
				label_url = "dhl/dhl-request-label-v4.asp?single=yes&newlabel=yes&invoiceid=" & rsGetOrder.Fields.Item("ID").Value
				print_label_url = "dhl/dhl-print-labels.asp?request_amount=single&invoiceid=" & rsGetOrder.Fields.Item("ID").Value
				shipping_company = "DHL"
			end if
			if instr(rsGetOrder.Fields.Item("shipping_type").Value,"USPS") > 0 then
				label_url = "usps/usps-request-label.asp?single=yes&newlabel=yes&invoiceid=" & rsGetOrder.Fields.Item("ID").Value
				print_label_url = "usps/usps-print-single-label.asp?invoiceid=" & rsGetOrder.Fields.Item("ID").Value
				shipping_company = "USPS"
			end if
			%>
			<a class="btn btn-sm btn-secondary d-inline-block" style="<%= display_shipping_label %>" id="reprint-label" href="<%= print_label_url %>" target="_blank">Re-print <%= shipping_company %> label</a>
			<button id="request-label" class="btn btn-sm btn-secondary d-inline-block" data-url="<%= label_url %>" data-shipper="<%= shipping_company %>">Request <%= shipping_company %> label</button>
			<button id="return-label" class="btn btn-sm btn-secondary d-inline-block"  data-toggle="modal" data-target="#modal-return-label">Return label</button>
			</div>
			<% if rsGetOrder.Fields.Item("USPS_tracking").Value <> "" then %>
				<%if rsGetOrder("checkout_estimated_delivery_date") <> "" then %>
					<div class="mb-1">
						<%=FormatDateTime(rsGetOrder("checkout_estimated_delivery_date"),vbShortDate)%> - Estimated original delivery date given at checkout
					</div>
				<% end if %>
				<%if rsGetOrder("estimated_delivery_date") <> "" then %>
					<div class="mb-1">
						<%=FormatDateTime(rsGetOrder("estimated_delivery_date"),vbShortDate)%> - Estimated delivery date at time of DHL label request
					</div>
				<% end if %>
				<% if instr(rsGetOrder.Fields.Item("shipping_type").Value,"DHL") > 0 then %>
				<span id="tracking_arrow_down" class="usps_tracking btn btn-sm btn-secondary" data-url="../dhl/dhl-tracking.asp?tracking=">Hide Tracking Details</span>
				<% else %>
				<span id="tracking_arrow_down" class="usps_tracking btn btn-sm btn-secondary" data-url="../usps_tools/usps_tracking.asp?id=">Hide Tracking Details</span>
				<% end if %>
				<span id="tracking_arrow_up" class="usps_tracking btn btn-sm btn-secondary" style="display:none">Show Tracking Details</span>
			<% end if %>
			<div class="form-group ups-tracking">
				<label class="font-weight-bold" for="ups-tracking">UPS Tracking:</label>
				<input name="ups-tracking" class="form-control form-control-sm" data-column="UPS_tracking" data-friendly="UPS tracking" type="text" value="<%=(rsGetOrder.Fields.Item("UPS_tracking").Value)%>">
			</div>
			<button class="btn btn-sm btn-secondary ml-1 mr-0" id="btn-send-shipment-email" data-invoiceid="<%= rsGetOrder.Fields.Item("ID").Value %>">Send shipment email</button><span id="msg-send-shipment-email"></span>
	
			<div class="mt-3" id="tracking_display" <% if rsGetOrder.Fields.Item("USPS_tracking").Value = "" then %> style="display:none" <%End If%>>
				<% 
				if rsGetOrder("USPS_tracking") <> "" then
				if instr(rsGetOrder.Fields.Item("shipping_type").Value,"DHL") > 0 then %>
					<script>$('#tracking_display').load("/dhl/dhl-tracking.asp?tracking=<%=rsGetOrder.Fields.Item("USPS_tracking").Value%>");</script>
				<% else %>
					<script>$('#tracking_display').load("/usps_tools/usps_tracking.asp?id=<%=rsGetOrder.Fields.Item("USPS_tracking").Value%>");</script>
				<% end if 
				end if
				%>				
			</div>

</div>
	<div class="col-sm pl-4 small disable-fields"> 
			<div class="container w-100">
					<div class="row">
						<div class="col-8 p-0 h4">Billing information</div>
						<div class="col-4 p-0 text-right">
								<!--Open payment details modal -->
							<button type="button" class="btn btn-sm btn-secondary" data-toggle="modal" data-target="#modal-payment-details">
  							More...
						</button>
						</div>
					</div>
				</div>          
<div id="message-frm-terminal"></div>

<form id="frm-terminal">	
		<div class="form-group">	
			<label class="font-weight-bold" for="pay-method">Payment method</label>
			<select class="form-control form-control-sm" name="pay-method" id="pay-method" data-column="pay_method" data-friendly="Payment method">
				<option value="<%=(rsGetOrder.Fields.Item("pay_method").Value)%>" selected><%=(rsGetOrder.Fields.Item("pay_method").Value)%></option>
				<option value="Visa">Visa</option>
				<option value="Mastercard">Mastercard</option>
				<option value="PayPal">PayPal</option>
				<option value="American Express">American Express</option>
				<option value="Discover">Discover</option>
				<option value="Etsy">Etsy</option>
				<option value="Instagram">Instagram</option>
				<option value="Facebook">Facebook</option>
				<option value="Cash">Cash</option>
			</select>
		</div>
			
<% if rsGetOrder.Fields.Item("date_sent").Value < "5/31/2017" and rsGetOrder.Fields.Item("pay_method").Value = "PayPal" then %>
<br/>
<div class="alert alert-danger">
	This PayPal transaction needs to be refunded through the PayPal website.
</div>
<% end if ' if paypal from a prior date %>
<div class="form-inline">
		<% if rsGetOrder.Fields.Item("payment_profile_id").Value <> 0 then %>
		<div class="custom-control custom-radio mr-5">
			<input name="tender" id="charge_cim" type="radio" value="charge_cim" class="custom-control-input tender">
			<label class="custom-control-label" for="charge_cim">Charge</label>
			
		</div>
		<% end if %>
		<% if rsGetOrder.Fields.Item("pay_method").Value <> "Afterpay" and rsGetOrder.Fields.Item("pay_method").Value <> "Instagram" then %>
		<div class="custom-control custom-radio mr-5">
			<input name="tender" type="radio" id="refund_card" value="refund" class="custom-control-input tender">
			<label class="custom-control-label" for="refund_card">Refund</label>
		</div>
		<% end if %>
		<% if rsGetOrder.Fields.Item("pay_method").Value = "PayPal" then %>
		<div class="custom-control custom-radio">
			<input name="tender" type="radio" value="money-request" id="money-request" class="custom-control-input tender">
			<label class="custom-control-label" for="money-request">Money request</label>
		</div>
		<% end if %>
		<% if rsGetOrder.Fields.Item("pay_method").Value <> "PayPal" and rsGetOrder.Fields.Item("pay_method").Value <> "Afterpay" and rsGetOrder.Fields.Item("pay_method").Value <> "Instagram" then %>
		<div class="custom-control custom-radio">
			<input name="tender" type="radio" value="void" id="void_charge"  class="custom-control-input tender">
			<label class="custom-control-label" for="void_charge">Void</label>
		</div>
		<% end if %>
		<% if rsGetOrder.Fields.Item("pay_method").Value = "Afterpay" then %>
		<div class="custom-control custom-radio mr-5">
			<input name="afterpay_payments" type="radio" value="refund" id="afterpay_payments_refund"  class="custom-control-input">
			<label class="custom-control-label" for="afterpay_payments_refund">Refund</label>
		</div>
		<% end if %>
	</div>
		<div class="form-group mt-2">
			<input class="form-control form-control-sm charge_amount" name="amount" type="text" placeholder="Amount $">     
		</div>
		<% if rsGetOrder.Fields.Item("pay_method").Value <> "Afterpay" then %>
		<div class="form-group">
			<textarea class="form-control form-control-sm charge_description" name="description" placeholder="For what items / changes:" ></textarea>
		</div>
		<% end if %>
		<div class="form-group">
			<input  class="form-control form-control-sm" name="trans_id" type="text" value="<%=(rsGetOrder.Fields.Item("transactionID").Value)%>" placeholder="Transaction ID #" class="transaction_id" data-column="transactionID" data-friendly="Transaction ID">
		</div>

		<button class="btn btn-sm btn-secondary" type="submit">Submit</button>
        
		<input name="first" type="hidden" value="<%=(rsGetOrder.Fields.Item("customer_first").Value)%>">  
		<input name="last" type="hidden" value="<%=(rsGetOrder.Fields.Item("customer_last").Value)%>">  
		<input name="email" type="hidden" value="<%=(rsGetOrder.Fields.Item("email").Value)%>">  
		<input name="invoice" type="hidden" value="<%=(rsGetOrder.Fields.Item("ID").Value)%>">  
		<input name="card_number" id="card_number" type="hidden" value="<%= replace(strCardNumber, "X", "") %>">
		<input name="customerProfileId" type="hidden" value="<%=(rsGetOrder.Fields.Item("cim_id").Value)%>"> 
		<input name="customerPaymentProfileId" type="hidden" value="<%=(rsGetOrder.Fields.Item("payment_profile_id").Value)%>"> 
		<input name="customerShippingAddressId" type="hidden" value="<%=(rsGetOrder.Fields.Item("shipping_profile_id").Value)%>"> 
</form>

		<% if rsGetOrder.Fields.Item("customer_comments").Value <> "" then %>
			<div class="my-2 text-info font-weight-bold">Customer comments: <%=rsGetOrder.Fields.Item("customer_comments").Value %>
			</div>
		<% end if %>
		<div class="form-group">
			 <textarea class="form-control form-control-sm" name="public_notes" id="public_notes" data-column="item_description" data-friendly="Public notes" placeholder="PUBLIC notes to print on invoice"  rows="5"><%=(rsGetOrder.Fields.Item("item_description").Value)%></textarea>
		</div>
 <div class="notes-icons">
	 <button title="Order updated" class="btn btn-sm btn-secondary insert-notes" data-text="ORDER UPDATED"><i class="fa fa-check"></i></button>
	 <button title="Address updated" class="btn btn-sm btn-secondary insert-notes" data-text="ADDRESS UPDATED"><i class="fa fa-address-book"></i></button>
	 <button title="Shipping method updated" class="btn btn-sm btn-secondary insert-notes" data-text="SHIPPING METHOD UPDATED"><i class="fa fa-truck"></i></button>
 </div>
	</div><!-- row -->
	</div><!-- container -->	

<div class="mt-4 disable-fields">
	<button type="button" class="btn btn-sm btn-secondary" data-toggle="modal" data-target="#modal-add-product">
			Add product(s) to order
	  </button>
	  <button class="btn btn-secondary btn-sm ml-3  btn_returns" data-toggle="modal" data-target="#modal-returns">Returns</button>
	  <button class="btn btn-sm btn-secondary ml-3 mr-0 btn-update-reship-modal" data-toggle="modal" data-target="#modal-reship-items" data-invoiceid="<%= rsGetOrder.Fields.Item("ID").Value %>">Reship items</button>
	  <i class="fa fa-lg fa-information text-secondary pointer" data-toggle="modal" data-target="#modal-reship-info"></i>
</div>
<table class="table table-striped table-hover small mt-2 ajax-update disable-fields">
<thead class="thead-dark">
	<tr>
		<th scope="col">
			<button class="btn btn-sm btn-secondary expand" type="button" data-text="Collapse all">Expand all</button>
		</th>
		<th scope="col">&nbsp;</th>
		<th scope="col">In stock</th>
		<th scope="col">Qty</th>
		<th scope="col">Product</th>
		<th scope="col">Sec</th>
		<th scope="col">Loc</th>
		<th scope="col">Price</th>
		<th scope="col">Total</th>
		<th scope="col">Coupon</th>
		<th scope="col">Color Fee</th>
		<th scope="col">Notes</th>
		<th scope="col">&nbsp;</th>
		<th scope="col">&nbsp;</th>
	</tr>
</thead>
<tbody >
<% 
	LineItem = 0
	SumLineItem = 0
	copy_order_header = ""
	copy_order_details = ""
	copy_line_detail = ""
	copy_totals_line = ""
	copy_totals = ""
While NOT rsGetOrderItems.EOF 
	if rsGetOrderItems.Fields.Item("returned").Value = 1 OR rsGetOrderItems("item_problem") <> "0" then
		class_returned = " table-danger "
	else
		class_returned = ""
	end if

' Show red notice class for errors in review 
if rsGetOrderItems.Fields.Item("ErrorOnReview").Value = 1 then
	review_highlight = "table-danger"
	open_default = ""
else
	review_highlight = ""
	open_default = "style=""display:none"""
end if

	copy_anodization_fee = ""
	if rsGetOrderItems("anodization_fee") > 0 then 
		copy_anodization_fee = " + " & FormatCurrency(rsGetOrderItems("qty") * rsGetOrderItems("anodization_fee"), -1, -2, -0, -2) & " color add-on fee"
	end if


	copy_line_detail = rsGetOrderItems.Fields.Item("qty").Value & "&nbsp;&nbsp;|&nbsp;&nbsp;" & rsGetOrderItems.Fields.Item("title").Value & "&nbsp;" & rsGetOrderItems.Fields.Item("ProductDetail1").Value & "&nbsp;" & rsGetOrderItems.Fields.Item("Gauge").Value & "&nbsp;" & rsGetOrderItems.Fields.Item("Length").Value & "&nbsp;" & rsGetOrderItems.Fields.Item("PreOrder_Desc").Value & "&nbsp;" & rsGetOrderItems.Fields.Item("notes").Value &  "&nbsp;&nbsp;&nbsp;&nbsp;$" & FormatNumber(rsGetOrderItems.Fields.Item("item_price").Value * rsGetOrderItems.Fields.Item("qty").Value, -1, -2, -0, -2) & "&nbsp;" & copy_anodization_fee
						
	copy_order_details = copy_order_details & "&#10;" &  copy_line_detail
%>


	<tr class="show-less <%= class_returned %> <%= review_highlight %> detail-main-<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" id="tbody-<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>">

		<% if rsGetOrder.Fields.Item("total_preferred_discount").Value <> 0 or rsGetOrder.Fields.Item("total_coupon_discount").Value <> 0 then

			if NOT rsGetCouponDiscount.eof then
			
				if rsGetCouponDiscount.Fields.Item("DiscountPercent").Value <> "" AND rsGetOrderItems.Fields.Item("SaleExempt").Value = 0 then
					discount_price = FormatNumber((rsGetOrderItems.Fields.Item("item_price").Value - ((rsGetCouponDiscount.Fields.Item("DiscountPercent").Value / 100) * rsGetOrderItems.Fields.Item("item_price").Value)) * rsGetOrderItems.Fields.Item("qty").Value, -1, -2, -0, -2)
				else 
				discount_price = FormatNumber(rsGetOrderItems.Fields.Item("item_price").Value * rsGetOrderItems.Fields.Item("qty").Value, -1, -2, -0, -2)
				end if
			
			else
			
			discount_price = 0
			
			end if
		
		end if

		LineItem = rsGetOrderItems.Fields.Item("item_price").Value * rsGetOrderItems.Fields.Item("qty").Value
		%>
	
		<td>
			<a class="expand-one btn btn-sm btn-outline-secondary mr-3" data-id="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>"><i class="fa fa-angle-down"></i></a>
		</td>
		<td>
			<span class="btn btn-sm btn-danger p-1 delete_item" data-delete_id="<%=(rsGetOrderItems.Fields.Item("OrderDetailID").Value)%>" data-price="<%= FormatNumber(discount_price, -1, -2, -0, -2) %>" data-origprice="<%= FormatNumber(LineItem, -1, -2, -0, -2) %>" data-qty="<%= rsGetOrderItems.Fields.Item("qty").Value %>"><i class="fa fa-trash-alt"></i></span>
		</td>
		<td><%=(rsGetOrderItems.Fields.Item("stock_qty").Value)%></td>
		<td>
			<input class="form-control form-control-sm" style="width: 50px" name="qty_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" type="text" data-column="qty" data-friendly="Qty" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>"  data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" value="<%=(rsGetOrderItems.Fields.Item("qty").Value)%>">
		</td>
		<td>
		<a href="/productdetails.asp?ProductID=<%=(rsGetOrderItems.Fields.Item("ProductID").Value)%>" target="_blank"><img src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetOrderItems.Fields.Item("picture").Value)%>" class="float-left mr-2" style="width:40px;height:40px"/></a>
			<% if rsGetOrderItems.Fields.Item("SaleExempt").Value <> 0 then %>
<strong>SALE EXEMPT</strong>       
<% end if %>
        <a class="text-dark" href="product-edit.asp?ProductID=<%=(rsGetOrderItems.Fields.Item("ProductID").Value)%>&detailid=<%=(rsGetOrderItems.Fields.Item("ProductDetailID").Value)%>" target="_blank"><%=(rsGetOrderItems.Fields.Item("item_description").Value)%></a>
          
&nbsp;&nbsp;
<% if rsGetOrderItems.Fields.Item("addon_item").Value = 1 then %>
	<br/><strong>ADD ON ITEM <%= rsGetOrderItems.Fields.Item("date_added").Value %></strong>
<% end if %>
<% if rsGetOrderItems.Fields.Item("returned").Value = 1 then %>
	<br/><strong>RETURNED QTY <%= rsGetOrderItems.Fields.Item("returned_qty").Value %></strong>
<% end if %>
<% If rsGetOrderItems("anodization_fee") > 0 and rsGetOrderItems("PreOrder_Desc") <> "" then %>
<br/><span class="badge badge-info mt-1" style="font-size:1em"><%= rsGetOrderItems("PreOrder_Desc") %> &#151; Anodization service added</span>
<% end if %>
<% if InStr( 1, lcase(rsGetOrderItems.Fields.Item("title").Value), lcase("gift certificate"), vbTextCompare) then 

' Get gift cert information
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM TBLcredits WHERE invoice = ?" 
objCmd.Parameters.Append objCmd.CreateParameter("invoice", 200, 1, 50, rsGetOrder.Fields.Item("ID").Value)
Set rsGetGiftCert = objCmd.Execute()

if not rsGetGiftCert.eof then
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
	objCmd.CommandText = "SELECT * FROM dbo.TBL_Credits_UsedOn WHERE OriginalCreditID = ?" 
	objCmd.Parameters.Append objCmd.CreateParameter("creditit", 200, 1, 50, rsGetGiftCert.Fields.Item("ID").Value)
	Set rsCertUsedOn = objCmd.Execute
end if
%>
<br/><a href="giftcertificate.asp?ID=<%= rsGetOrder.Fields.Item("ID").Value %>&amp;amt=<%= rsGetOrderItems.Fields.Item("item_price").Value %>&send=no" target="_blank">Re-send gift certificate e-mail</a>
<% if not rsGetGiftCert.eof then %>
<br/><br/>
<strong>Recipient name:</strong> <%= rsGetGiftCert.Fields.Item("rec_name").Value %><br/>
<strong>Recipient email:</strong> <%= rsGetGiftCert.Fields.Item("rec_email").Value %><br/>
<strong>Code:</strong> <%= rsGetGiftCert.Fields.Item("code").Value %><br/>
<strong>Amount:</strong> <%= FormatCurrency(rsGetGiftCert.Fields.Item("amount").Value, -1, -2, -0, -2) %>
<% 
 if not rsCertUsedOn.eof then %>
 <br/>
 <br/>
<strong>USED ON:</strong><br/>
<% while not rsCertUsedOn.eof %>
	<a href="invoice.asp?ID=<%= rsCertUsedOn.Fields.Item("InvoiceUsedOn").Value %>" target="_blank"><%= rsCertUsedOn.Fields.Item("InvoiceUsedOn").Value %></a><br/>
<% 
rsCertUsedOn.movenext()
wend

end if ' if gift cert not used on any order
end if ' display gift cert info
 end if 'if gift cert is found 
 %>
<% if InStr( 1, (rsGetOrderItems.Fields.Item("title").Value), "CUSTOM ORDER", vbTextCompare) then

if rsGetOrderItems.Fields.Item("item_ordered").Value = 1 then
	var_ordered_status = "alert-success"
	if rsGetOrderItems.Fields.Item("item_ordered_date").Value <> "" then
		var_ordered_information = "<span class='font-weight-bold'>Ordered on " & FormatDateTime(rsGetOrderItems.Fields.Item("item_ordered_date").Value, 2) & " and " & rsGetOrderItems.Fields.Item("preorder_timeframes").Value & "</span>"
	end if
else
	var_ordered_status = "alert-danger"
	var_ordered_information = "" 
end if

if rsGetOrderItems.Fields.Item("item_received").Value = 1 then
	var_received_status = "alert-success"
	if rsGetOrderItems.Fields.Item("item_received_date").Value <> "" then
		var_received_information = "<span class='font-weight-bold'>Received on " & FormatDateTime(rsGetOrderItems.Fields.Item("item_received_date").Value, 2) & "</span>"
	end if
else
	var_received_status = "alert-danger"
	var_received_information = ""
end if
%>
<textarea  class="mt-2 form-control form-control-sm" name="preorder_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" type="text" data-column="PreOrder_Desc" data-friendly="Custom item specs"  data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>"><% If isNull(rsGetOrderItems.Fields.Item("PreOrder_Desc").Value) then %><% else %><%= Server.HTMLEncode(rsGetOrderItems.Fields.Item("PreOrder_Desc").Value) %><% end if %></textarea>

 <label class="m-0 mt-2" for="ordered_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>">Ordered: <%= var_ordered_information %></label> 
          <select  class="form-control form-control-sm status <%= var_ordered_status %>" name="ordered_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" type="text" data-column="item_ordered" data-friendly="Item ordered" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>">
            <option <% if (rsgetorderitems.fields.item("item_ordered").value) <> 1 then %>value="0"<% else %>value="1"<% end if%> selected>
              <% if (rsGetOrderItems.Fields.Item("item_ordered").Value) <> 1 then %>
              no
              <% else %>
              YES
              <% end if%>
            </option>
            <option value="0">no</option>
            <option value="1">yes</option>
		  </select>
		  
		  <label class="m-0 mt-2" for="received_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>">Received: <%= var_received_information %></label>
           
          <select  class="form-control form-control-sm status <%= var_received_status %>" name="received_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" type="text" data-column="item_received" data-friendly="Item received" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>">
            <option <% if (rsGetOrderItems.Fields.Item("item_received").Value) <> 1 then %>value="0"<% else %>value="1"<% end if%> selected><% if (rsGetOrderItems.Fields.Item("item_received").Value) <> 1 then %>no<% else %>YES<% end if%></option>
			<option value="0">no</option>
            <option value="1">yes</option>
          </select>
          <% else %>
          <% end if %>

		</td>
		<td><%=(rsGetOrderItems.Fields.Item("ID_Description").Value)%></td>
		<td>
			<%=(rsGetOrderItems.Fields.Item("location").Value)%>
			<% if rsGetOrderItems.Fields.Item("BinNumber_Detail").Value <> 0 then %>
				(BIN <%=(rsGetOrderItems.Fields.Item("BinNumber_Detail").Value)%>)
			<% end if %>
		</td>
		<td>
			<input class="form-control form-control-sm" style="width: 75px" name="retail_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" type="text" value="<%= FormatNumber(rsGetOrderItems.Fields.Item("item_price").Value, -1, -2, -0, -2) %>" data-column="item_price" data-friendly="Price" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>">
		</td>
		<td>
			<%= FormatCurrency(LineItem, -1, -2, -0, -2) %>
		</td>
		<% if rsGetOrder.Fields.Item("total_preferred_discount").Value <> 0 or rsGetOrder.Fields.Item("total_coupon_discount").Value <> 0 then
				' if a coupon was used pass the discounted price to the BO page, if not use the regular retail price
				bo_price = discount_price
		%>
		<td>
			<span class="alert alert-danger px-2 py-1"><%= FormatCurrency(discount_price, -1, -2, -0, -2) %></span>
		</td>
		<% else 
			bo_price = LineItem
		%>
		<td>&nbsp;</td>
		<% end if %>
		<td>
			<%= FormatCurrency(rsGetOrderItems("qty") * rsGetOrderItems("anodization_fee"), -1, -2, -0, -2) %>
		</td>
		<td>
			<input class="form-control form-control-sm" name="item_notes_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" type="text" value="<%= rsGetOrderItems.Fields.Item("notes").Value %>" data-column="notes" data-friendly="Item notes" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>">
		</td>
		<td>
			<% if rsGetOrderItems.Fields.Item("backorder").Value = 0 then
				bo_visibility = ""
				on_bo_visibility = " style='display:none' "
			else
				bo_visibility = " style='display:none' "
				on_bo_visibility = ""
			end if
			%>
			<button class="btn btn-sm btn-info font-weight-bold p-0 px-1 border-0 bo_blue_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" <%= bo_visibility %>  type="button" id="btn-update-bo-modal" data-toggle="modal" data-target="#modal-submit-backorder" data-itemid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-qty="<%= rsGetOrderItems.Fields.Item("stock_qty").Value %>" data-title="<%= Server.HTMLEncode(rsGetOrderItems.Fields.Item("item_description").Value) %>">BO</button>
			<button class="btn btn-sm btn-warning font-weight-bold p-0 px-1 border-0 process-bo bo_orange_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>"  <%= on_bo_visibility %> type="button" data-id="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-price="<%= FormatNumber(bo_price, -1, -2, -0, -2) %>" data-origprice="<%= FormatNumber(LineItem, -1, -2, -0, -2) %>" data-qty="<%= rsGetOrderItems.Fields.Item("qty").Value %>" data-detailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-qty_instock="<%= rsGetOrderItems.Fields.Item("stock_qty").Value %>"  data-toggle="modal" data-target="#modal-backorder">On BO</button>

		</td>
		<td style="white-space: nowrap">
			<span class="input_move" name="input_move_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>"><span><button class="btn btn-sm p-0 px-1 border-0  btn-secondary d-inline-block font-weight-bold copyid" name="copy_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-id="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-qty_instock="<%= rsGetOrderItems.Fields.Item("stock_qty").Value %>" data-qty="<%= rsGetOrderItems.Fields.Item("qty").Value %>">C</button>
			<button class=" btn btn-sm btn-secondary font-weight-bold p-0 px-1 border-0  d-inline-block moveid" name="move_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-id="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>">M</button>
		</td>
	</tr>

	<tr class="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %> expanded-details" <%= open_default %>>
		<td colspan="14">
			
			<% 'if rsGetOrderItems.Fields.Item("ErrorReportDate").Value  <> "" OR Request.Querystring("ReportError") = "yes" then %>
			<div class="form-inline">
			<label>Error type:</label>
			<select name="ErrorType_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-column="item_problem" data-friendly="Error type" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" class="form-control form-control-sm ml-2 mr-4">
			<option value="<% If (rsGetOrderItems.Fields.Item("item_problem").Value) <> "" then %><%=(rsGetOrderItems.Fields.Item("item_problem").Value)%>" selected<% end if %>><%=(rsGetOrderItems.Fields.Item("item_problem").Value)%></option>
			<option value="0" <% If rsGetOrderItems.Fields.Item("item_problem").Value = "" or rsGetOrderItems.Fields.Item("item_problem").Value = "0" OR isNull(rsGetOrderItems.Fields.Item("item_problem").Value) then %>selected<% end if %>>No errors</option>
			<option value="Broken">Broken</option>
			<option value="Missing">Missing</option>
			<option value="Wrong">Wrong</option>
			<option value="Mis-matched">Mis-matched</option>
			<option value="Flip-flop">Flip-flop</option>
			<option value="Misc">Misc</option>
			</select>
		
			<label>Qty to reship:</label>
			<input name="qty_missing_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" type="text" class="form-control form-control-sm ml-2 mr-4 qty"  style="width:70px" value="<% = (rsGetOrderItems.Fields.Item("ErrorQtyMissing").Value) %>" data-column="ErrorQtyMissing" data-friendly="Qty to reship" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>">
		
			
			<select name="sealedBag_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-column="ErrorSealedBag" data-friendly="Sealed bag error" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" class="form-control form-control-sm mr-4">
				<option value="N" <% If (rsGetOrderItems.Fields.Item("ErrorSealedBag").Value) = "N" OR (rsGetOrderItems.Fields.Item("ErrorSealedBag").Value) = "" then %>selected<% end if %>>
					Sealed ok
				</option>
				<option value="Y" <% If (rsGetOrderItems.Fields.Item("ErrorSealedBag").Value) = "Y" then %>selected<% end if %>>
					Not sealed ok
				</option>
			</select>
			
		
			
			<select name="conservedBags_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-column="ErrorConserveBags" data-friendly="Conserved bags error" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" class="form-control form-control-sm">
				<option value="Y" <% If (rsGetOrderItems.Fields.Item("ErrorConserveBags").Value) = "Y" then %>selected<% end if %>>
					Error conserving bags
				</option>
				<option value="N" <% If (rsGetOrderItems.Fields.Item("ErrorConserveBags").Value) = "N" then %>selected<% end if %>>
					Conserved bags properly
				</option>
			</select>
		</div>
		
			<%' if rsGetOrderItems.Fields.Item("ErrorOnReview").Value = 1 then %>
			<div class="my-2">
			Do NOT show on review error page?
			<div class="custom-control custom-radio d-inline-block mx-3">
			<input class="custom-control-input" type="radio" name="reviewError_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" id="yes_reviewError_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-column="ErrorOnReview" data-friendly="Error on review" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" value="0" <% if rsGetOrderItems.Fields.Item("ErrorOnReview").Value = 0 then %>checked<% end if %>>
			<label class="custom-control-label" for="yes_reviewError_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>">Yes</label>
			</div>
			<div class="custom-control custom-radio d-inline-block">
			<input class="custom-control-input" type="radio" name="reviewError_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" id="no_reviewError_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-column="ErrorOnReview" data-friendly="Error on review" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" value="1" <% if rsGetOrderItems.Fields.Item("ErrorOnReview").Value = 1 then %>checked<% end if %>>
			<label class="custom-control-label" for="no_reviewError_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>">No</label>
			</div>
		</div>
			<%' end if %>
			
			<div class="form-inline">
			<label>Date reported:</label>
				<input class="form-control form-control-sm ml-2" name="errorReportDate_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-column="ErrorReportDate" data-friendly="Reported error date" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" type="text" value="<%=(rsGetOrderItems.Fields.Item("ErrorReportDate").Value)%>">
			</div>

			<div class="form-group mt-2">
				<textarea class="form-control form-control-sm" name="ErrorDescription_<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" cols="60" rows="3" data-column="ErrorDescription" data-friendly="Error description" data-detailid="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetOrderItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetOrderItems.Fields.Item("ProductID").Value %>" placeholder="&nbsp;Description of error"><%=(rsGetOrderItems.Fields.Item("ErrorDescription").Value)%></textarea>
			</div>
			<% 'end if %>
		</td>
	</tr>
<%
	SumLineItem = SumLineItem + LineItem
	sum_anodization_fees = sum_anodization_fees + rsGetOrderItems("qty") * rsGetOrderItems("anodization_fee")
rsGetOrderItems.MoveNext()
Wend
	InvoiceTotal = SumLineItem + (rsGetOrder.Fields.Item("shipping_rate").Value) - (rsGetOrder.Fields.Item("coupon_amt").Value)
copy_totals = "Subtotal: &nbsp;&nbsp;&nbsp;" & FormatCurrency(SumLineItem, -1, -2, -0, -2) & "&#10;"

%>
</tbody>
</table>

<table class="table table-hover small disable-fields">
		<tbody>
			<tr>
					<td class="border-0" style="width:60%">
						<button class="btn btn-sm btn-secondary mr-4" id="duplicate_order" data-invoiceid="<%= rsGetOrder("ID") %>" data-email="<%= rsGetOrder("email") %>">Duplicate order <span id="msg-duplicate-order"></span></button>	
						<button class="btn btn-sm btn-secondary mr-4" id="create_new_order">Create new empty order</button>
							<span class="move-copy-productid form-inline" style="display:none">
							<div class="mt-2">
							<span id="move-copy-text"></span> to invoice # <input class="form-control form-control-sm" type="text" size="20" name="toggle-productid" placeholder="Invoice # or new">
							</span></div>
						</td>	
				<td style="width:30%">
					Coupon code
						</td>	
				<td style="width:10%">
						<input name="coupon_code" data-column="coupon_code" data-friendly="Coupon code" type="text" value="<%= rsGetOrder.Fields.Item("coupon_code").Value %>" placeholder="Coupon code" class="form-control form-control-sm mr-2">
					</td>
			</tr>
<%
' Array for invoice totals
Dim arrTotals(2,6) 

'arrTotals(col,row)
arrTotals(0,0) = "10% preferred discount" 
arrTotals(1,0) = "total_preferred_discount" 
total_preferred_discount = rsGetOrder.Fields.Item("total_preferred_discount").Value
arrTotals(2,0) = "&#8722;"
arrTotals(0,1) = "Coupon discount" 
arrTotals(1,1) = "total_coupon_discount" 
total_coupon_discount = rsGetOrder.Fields.Item("total_coupon_discount").Value
arrTotals(2,1) = "&#8722;" 
arrTotals(0,2) = "Tax (" & rsGetOrder.Fields.Item("combined_tax_rate").Value * 100 & "%)"
arrTotals(1,2) = "total_sales_tax" 
total_sales_tax = rsGetOrder.Fields.Item("total_sales_tax").Value
arrTotals(2,2) = "&nbsp;&nbsp;"
arrTotals(0,3) = "Gift certificate" 
arrTotals(1,3) = "total_gift_cert"
total_gift_cert = rsGetOrder.Fields.Item("total_gift_cert").Value 
arrTotals(2,3) = "&#8722;"
arrTotals(0,4) = "Free gift (USE NOW) credits" 
arrTotals(1,4) = "total_free_credits" 
total_free_credits = rsGetOrder.Fields.Item("total_free_credits").Value
arrTotals(2,4) = "&#8722;"
arrTotals(0,5) = "Store account credit" 
arrTotals(1,5) = "total_store_credit"
total_store_credit = rsGetOrder.Fields.Item("total_store_credit").Value
arrTotals(2,5) = "&#8722;"
arrTotals(0,6) = "Order returns" 
arrTotals(1,6) = "total_returns"
total_returns = rsGetOrder.Fields.Item("total_returns").Value
arrTotals(2,6) = "&#8722;"


For i = 0 to UBound(arrTotals, 2) 

'	if i <=1 or rsGetOrder.Fields.Item(arrTotals(1,i)).Value <> 0 or (rsGetOrder.Fields.Item("state").Value = "TX")  then
%>
<tr>
		<td class="border-0">
				<% if arrTotals(1,i) = "total_preferred_discount" then %>
				<div class="move-copy-productid" style="display:none">
						<div class="custom-control custom-checkbox d-inline-block mr-4 ">
						<input class="custom-control-input" name="ReturnMailer" id="ReturnMailer" type="checkbox" value="Yes">
						<label class="custom-control-label" for="ReturnMailer">Send a return mailer?</label>
						</div>
						<div class="custom-control custom-checkbox d-inline-block">
						<input class="custom-control-input" name="reship-returned" type="checkbox" id="reship-returned" value="Yes">
						<label class="custom-control-label" for="reship-returned">Ship returned order?</label>
						</div>
					</div>
						<% end if %>
			<% if arrTotals(1,i) = "total_returns" then %>
				<button class="btn btn-sm btn-secondary mr-3 update_inventory" data-type="deduct">Deduct quantities</button>
				<button class="btn btn-sm btn-secondary update_inventory" data-type="add">Add quantities</button>
				<span id="confirm_inv_updates" class="alert alert-success p-2 ml-3 font-weight-bold" style="display:none">Inventory has been updated</span>
			<% end if %>
		</td>
		<td class="form-inline">
		
				<%= arrTotals(0,i) %>
		
			<% if arrTotals(1,i) = "total_gift_cert" then 
			
				Set objCmd = Server.CreateObject ("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "SELECT TBL_Credits_UsedOn.InvoiceUsedOn, TBLcredits.invoice, TBLcredits.code FROM TBL_Credits_UsedOn INNER JOIN TBLcredits ON TBL_Credits_UsedOn.OriginalCreditID = TBLcredits.ID WHERE TBL_Credits_UsedOn.InvoiceUsedOn = ?" 
				objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12,rsGetOrder.Fields.Item("ID").Value))
				Set rsGetGiftCertInfo = objCmd.Execute()
				
				if not rsGetGiftCertInfo.eof then
				var_gift_cert_code_used = rsGetGiftCertInfo.Fields.Item("code").Value
			%>
				
				&nbsp;<%= rsGetGiftCertInfo.Fields.Item("code").Value %>
				
				# <%= rsGetGiftCertInfo.Fields.Item("invoice").Value %>
				
			<%	end if '  if not rsGetGiftCertInfo.eof
			end if %>
		</td>
		<td>
				<input class="form-control form-control-sm" name="<%= arrTotals(1,i) %>" type="text" value="<%= rsGetOrder.Fields.Item(arrTotals(1,i)).Value %>" data-column="<%= arrTotals(1,i) %>" data-friendly="<%= arrTotals(0,i) %>">
		</td>
	</tr>
<% 
'	end if ' if i > 2 or values not 0
if FormatNumber(rsGetOrder.Fields.Item(arrTotals(1,i)).Value) > 0 then
	copy_totals_line = arrTotals(0,i) & ":&nbsp;&nbsp;&nbsp;" & var_minus & FormatCurrency(rsGetOrder.Fields.Item(arrTotals(1,i)).Value, -1, -2, -0, -2) & "&#10;"
	copy_totals = copy_totals & "" & copy_totals_line
end if

next ' loop through totals array

InvoiceTotal = (SumLineItem + sum_anodization_fees - total_preferred_discount - total_coupon_discount - total_free_credits + rsGetOrder.Fields.Item("shipping_rate").Value + total_sales_tax - total_store_credit - total_gift_cert - total_returns)

copy_totals = copy_totals & "Shipping:&nbsp;&nbsp;&nbsp;" & FormatCurrency(rsGetOrder.Fields.Item("shipping_rate").Value, -1, -2, -0, -2) & "&#10;" & "&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#10;TOTAL:&nbsp;&nbsp;&nbsp;" & FormatCurrency(InvoiceTotal, -1, -2, -0, -2) & ""
copy_order_header = "Invoice # " & rsGetOrder.Fields.Item("ID").Value & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsGetOrder.Fields.Item("shipped").Value & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsGetOrder.Fields.Item("date_sent").Value & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsGetOrder.Fields.Item("shipping_type").Value

%>
</tbody>
</table>
<button class="btn btn-sm btn-secondary d-inline-block" id="copy-order" data-clipboard-text="<%= copy_order_header %>&#10;<%= replace(copy_order_details, """", " inch") %>&#10;&#10;<%= copy_totals %>"> <i class="fa fa-content-copy"></i> Copy order to clipboard</button>
</div><!-- main container -->

<div style="height:100px"></div>
<% 
else
%>
<h3 class="p-3">Invoice # <%= var_invoiceid %> not found.</h3>
<%
end if ' if not rsGetOrder.eof then 
%>

<div class="fixed-bottom bg-dark text-light text-center pt-1">
	<h4>
		Subtotal: $<%= SumLineItem %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		Shipping: <% if not rsGetOrder.eof then Response.Write FormatCurrency(rsGetOrder.Fields.Item("shipping_rate").Value, -1, -2, -0, -2) %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		GRAND TOTAL: $<span id="invoice-total"><% if InvoiceTotal < 0 then %>0<% else %><%= FormatNumber(InvoiceTotal, -1, -2, -0, -2) %><% end if %></span>
	</h4>
</div>

<!-- add product to order -->
<div class="modal fade small" id="modal-add-product" tabindex="-1" role="dialog"  aria-labelledby="modal-add-product" >
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title">Add Product(s) to Order</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body">
								
				<div id="td_search">
					<div class="form-group">
						<label>Product #</label>
						<input class="form-control form-control-sm" type="text" id="frm_search_product" placeholder="Product # to add / search">
					</div>
					<div id="search_results"></div>
					<div id="show_frm_add" style="display:none">
							<div class="container w-100 mt-3">
									<div class="row">
										<div class="col-6">
												<div class="form-group">
														<label class="font-weight-bold">Qty to add:</label>
														<input class="form-control form-control-sm" type="text" id="frm_qty" placeholder="Qty to add">
													</div>
										</div>
										<div class="col-6">
												<div class="form-group">
														<label class="font-weight-bold">Item price:</label>
														<input class="form-control form-control-sm" type="text" id="frm_price" placeholder="Item price" value="">
													</div>
										</div>
									</div>
								</div> 

						<input type="hidden" id="frm_detailid" value="">
						<span class="btn btn-sm btn-primary" id="btn_add_item">Add item</span>
					</div>
				</div>
			</div>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		  </div>
		</div>
	  </div>

<!-- Process backorder Modal -->
	<div class="modal fade" id="modal-backorder" tabindex="-1" role="dialog"  aria-labelledby="modal-backorder" >
		<div class="modal-dialog modal-dialog-scrollable modal-lg" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title">Process Backorder</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body small">
				<div class="load-bo"></div>
			</div>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		  </div>
		</div>
</div>
<!-- End Process backorder Modal -->

<!-- Process backorder Modal -->
	<div class="modal fade" id="modal-submit-backorder" tabindex="-1" role="dialog"  aria-labelledby="modal-submit-backorder" >
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title">Submit Backorder</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body small">
				<div id="new-bo-message"></div>
				<!--#include file="invoices/inc-submit-backorder.asp"-->
			</div>
			<div class="modal-footer">
				<button type="button" class="btn btn-primary" id="btn-submit-bo" data-itemid="">Submit</button>
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		  </div>
		</div>
</div>
<!-- End Process backorder Modal -->

<!-- Start RETURNS Modal -->
	<div class="modal fade" id="modal-returns" tabindex="-1" role="dialog"  aria-labelledby="modal-returns" >
		<div class="modal-dialog modal-lg modal-dialog-scrollable" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title">Process Returns</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body small">
				<button class="btn btn-sm btn-secondary btn-show-target return-agenda" data-show="return-undelivered" data-agenda="undelivered">Package came back notification</button>
				<button class="btn btn-sm btn-secondary btn-show-target return-agenda" data-show2="return-items-list" data-agenda="return-items" id="btn-return-items">Refund returned item(s)</button>
				<form class="return-hide return-items-list mt-4" style="display:none" id="form-returned-items-selection">
					<input type="hidden" name="invoiceid" value="<%= rsGetOrder.Fields.Item("id").Value %>">
					<input type="hidden" name="returns-ccrefund" id="returns-ccrefund" value="0">
					<input type="hidden" name="returns-storecredit-due" id="returns-storecredit-due" value="0">
					<input type="hidden" name="returns-giftcert-due" id="returns-giftcert-due" value="0">
					<input type="hidden" name="returns-sales-tax" id="returns-sales-tax" value="0">
					<input type="hidden" name="var_gift_cert_code_used" id="var_gift_cert_code_used" value="<%= var_gift_cert_code_used %>">
					<input name="returns_card_number" id="returns_card_number" type="hidden" value="<%= replace(strCardNumber, "X", "") %>">
					<input name="returns-calculation" id="returns-calculation" type="hidden">
					<% if rsGetOrder.Fields.Item("customer_id").Value <> 0 then %>
					<div class="custom-control custom-checkbox d-inline-block">
						<input type="checkbox" class="custom-control-input" name="store-credit-only" id="store-credit-only">
						<label class="custom-control-label" for="store-credit-only">Refund to store credit only</label>
					  </div>
					  <% end if %>
					  <div class="custom-control custom-checkbox d-inline-block ml-3">
						<input type="checkbox" class="custom-control-input" name="preorder-restock-fee" id="preorder-restock-fee">
						<label class="custom-control-label" for="preorder-restock-fee">15% custom item restock fee</label>
					  </div>

					<table class="table table-sm">
						<thead class="thead-dark">
							<tr>
								<th class="h5" colspan="4">TOTAL $<span id="returns-total"></span>
									<button type="button" class="ml-5 btn btn-sm btn-primary" id="btn-return-calculate">Calculate</button>
									<% if rsGetOrder.Fields.Item("shipping_rate").Value > 0 then %>
									<span class="btn-group-toggle ml-3" data-toggle="buttons">
											<label class="btn btn-sm btn-secondary" id="btn-refund-shipping">
											  <input type="checkbox" autocomplete="off" value="1" name="refund-shipping"><i class="fa fa-check mr-2" id="icon-toggle-shipping-refund" style="display:none"></i> Refund $<%= rsGetOrder.Fields.Item("shipping_rate").Value %> shipping
											</label>
										</span>
										<% end if %>
										<input type="text" class="form-control form-control-sm d-inline w-auto ml-4 bg-light" placeholder="Additional amount" name="additional_amount">
								</th>
							</tr>
						</thead>
						<tr class="table-light small">
								<td><i class="fa fa-plus-circle btn btn-sm btn-secondary" id="btn-return-selectall"></i></td>
								<td class="h6">Qty</td>
								<td class="h6">Item</td>
								<td class="h6">Price</td>
							</tr>
						<%
						if rsGetOrderItems.EOF and ( Not rsGetOrderItems.BOF) then
						rsGetOrderItems.MoveFirst()
						end if
						
						While NOT rsGetOrderItems.EOF 
						%>
						<tr>
							<td>
								<i class="fa fa-times-circle return-check btn btn-sm btn-danger" data-id="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>"></i>
							</td>
						<td>
							<input class="form-control form-control-sm return-qty" style="width: 40px" name="<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" id="return-id-<%= rsGetOrderItems.Fields.Item("OrderDetailID").Value %>" value="<%= rsGetOrderItems.Fields.Item("qty").Value %>" disabled>
						</td>
						<td class="small">
							<img src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetOrderItems.Fields.Item("picture").Value)%>" class="float-left mr-2" style="width:40px;height:40px"/>
							<%=(rsGetOrderItems.Fields.Item("item_description").Value)%>
						</td>
						<td class="small">
								$<%= FormatNumber(rsGetOrderItems.Fields.Item("item_price").Value * rsGetOrderItems.Fields.Item("qty").Value) %>
					</td>
						</tr>
						<% rsGetOrderItems.MoveNext()
						Wend
						%>
					</table>
				<div class="form-group">
						<label class="font-weight-bold">Comments:</label>
						<textarea class="form-control form-control-sm"   name="return-extra-comments" rows="3"></textarea>
					</div>
				</form>

				<form class="return-undelivered return-hide" style="display:none" id="frm-undeliverable" name="frm-undeliverable">
						<h6 class="mt-3">Why did the package come back?</h6>
						<div class="form-group">
							<select class="form-control form-control-sm" id="undeliverable-reason">
								<option>Select reason...</option>
								<option value="Unclaimed">Unclaimed</option>
								<option value="No mail receptacle available">No mail receptacle available</option>
								<option value="Attempted not known">Attempted not known</option>
								<option value="No apartment # or suite #">No apartment # or suite #</option>
								<option value="Undeliverable address">Undeliverable address</option>
								<option value="Moved and left no forwarding address">Moved and left no forwarding address</option>
								<option value="No reason given">No reason given</option>
								<option value="Damaged">Damaged</option>
								<option value="Other">Other</option>
							</select>
						</div>
						<div class="form-group return-hide" id="group-undeliverable-other" style="display:none">
							<label class="font-weight-bold">Other reason:</label>
							<textarea class="form-control form-control-sm"  id="undeliverable-reason-other" rows="3"></textarea>
						</div>
					  </form>
					  <div id="message-returns"></div>
			</div>
			<div class="modal-footer">
				<button type="button" class="btn btn-primary return-hide" style="display:none" id="submit-return" data-agenda="" data-invoiceid="<%= rsGetOrder.Fields.Item("ID").Value %>">Submit</button>
			  <button type="button" class="btn btn-secondary" id="btn-returns-close" data-dismiss="modal">Close</button>
			</div>
		  </div>
		</div>
</div>
<!-- End Returns Modal -->

<!-- Modal to load items from invoice to approve/deny -->
<div class="modal fade" id="modal-reship-items" tabindex="-1" role="dialog"  aria-labelledby="modal-reship-items" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Reship items</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
			<div id="message-reship-status"></div>
			<div id="load-reship-items"></div>
		</div>
		<div class="modal-footer">
			<div class="d-inline-block text-left w-50">
				<button type="button" class="btn btn-sm btn-danger btn-reship-items" data-agenda="deny" id="btn-reship-deny" data-invoiceid="">Deny reship</button>
			</div>
			<div class="d-inline-block text-right w-50">
				<button type="button" class="btn btn-sm btn-primary btn-reship-items" data-agenda="approve" id="btn-reship-approve" data-invoiceid="">Reship items</button>
			</div>
		</div>
	  </div>
	</div>
</div>
<!-- End Modal to load items from invoice to approve/deny -->


<!-- Modal to learn about how to use the reship items feature -->
<div class="modal fade" id="modal-reship-info" tabindex="-1" role="dialog"  aria-labelledby="modal-reship-info" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Reship items information</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
			<h6 class="my-1">Reviewing the items</h6>
			When you click the button it brings up a window with all the items that are currently set to order error review for that order that a customer submitted. But let's say a customer sent in the issue via e-mail. In order to use this feature you'll need to expand the items with issues from the invoice page and set them to "Do not show on review error page = NO".
			<br/><br/>
	  		Once the items show up in the window look the info over. The qty box allows you to change the items that will be reshipped. This is pulling the data that the customer submitted to us. It will overwrite it once you update the qty #.
			
			  <h6 class="mt-3 mb-1">Approving the reship</h6>
				
				
				
				Once you have all the qty #'s correct you can then Approve the reship. This *should* do everything. It will:
				<ul>
				<li>Remove the items back off review error</li>
				
				<li>Write all the necessary notes on the invoice level and item level</li>
				
				<li>Both USA and International will go DHL basic mail</li>
				
				<li>Send the customer an email with an entire breakdown of what is being shipped, credited, etc. There's no need to send the customer an email.</li>
				
				<li>If items are out of stock it will see if that customer is registered and give them a store credit</li>
				
				<li>If items are out of stock and the customer is NOT registered it will generate them a gift certificate</li>
				
				<li>If giving a refund/gift cert it will calculate the coupon discount on the item and then also add the tax back in so that it all gets correctly refunded</li>
			</ul>
				
				<h6 class="mb-1">Denying the reship</h6>
				
				Currently, this does nothing. No email will be sent. I may build something out into this area in the future, or get rid of it and come up with another area to address other issues. You can click anywhere on the page or the X to close out the window which is the same as a 'deny'.
		</div>
	  </div>
	</div>
</div>
<!-- End Modal to learnabout how to use the reship items feature  -->

<!-- RETURN LABEL MODAL -->
<div class="modal fade" id="modal-return-label" tabindex="-1" role="dialog"  aria-labelledby="modal-return-label" >
	<div class="modal-dialog modal-xl" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Return Labels</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
			We are only charged for the label once the customer drops the package off and the label is scanned into the system.
			<hr/>
			<h5>Domestic return labels</h5>
			Visit <a class="text-info font-weight-bold" href="https://portal.dhlecs.com/login.cfm" target="_blank">https://portal.dhlecs.com</a> to create a return USPS First Class Mail label.
			<br>
			Username: amanda.bunch  |  Password: 9U27uazrbh6Y!e4
			<ul>
				<li>Once you login to the DHL portal click Returns (on the top navigation) > Create a return label</li>
				<li>On Step 1 put in the customers address as the Return From.</li>
				<li>On Step 2 you can leave all of the optional fields blank.</li>
				<li>On Step 3 select the weight. For the from name and email use Bodyartforms and service@bodyartforms.com. Fill out the customers name and email. You can leave the optional fields blank.</li>
				<li>Customers will be e-mailed a PDF file with the label that they can print out and use.</li>
				<li>Customers can drop the package off at their mailbox or any location that has a USPS drop box.</li>
			</ul>
			<hr/>
			<h5>Canadian return labels</h5>
			<a href="https://www.canadapost.ca/information/app/prse/label?policyId=PR407013&LOCALE=en" target="_blank">Click here</a> to go to the Canada Post website to create a return label
			<p>
				Please note that you have to submit the <u>"GM + Invoice #"</u> (Example: GM123456) number as the reference number (which is DHLs Outbound Tracking # or GM Number). Without it they may not be able to process the parcel for return to US. Customers can drop their return shipment into any Canada Post box.
				<ul>
					<li>For the address field, fill in the CUSTOMERS address. DHL will know to ship the package back to us.</li>
					<li>Will take 2-3 weeks to get back to us</li>
					<li>Has to be within 60 days of the original print label date</li>
				</ul>
			</p>
		</div>
		<div class="modal-footer">
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		</div>
	  </div>
	</div>
  </div>

<!-- Show more billing information -->
<div class="modal fade" id="modal-payment-details" tabindex="-1" role="dialog"  aria-labelledby="modal-payment-details" >
	<div class="modal-dialog" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Payment Details</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
			<% If rsGetOrder.Fields.Item("pay_method").Value <> "PayPal" AND rsGetOrder.Fields.Item("pay_method").Value <> "Money order" AND rsGetOrder.Fields.Item("pay_method").Value <> "Cash"  AND rsGetOrder.Fields.Item("pay_method").Value <> "Instagram"  AND rsGetOrder.Fields.Item("pay_method").Value <> "Afterpay" then %>
			<strong><u>Address verification:</u></strong><br/>
			<strong>Card #</strong> <%= replace(strCardNumber, "X", "") %><br>
			<strong>AVS &#8213;</strong> <%= str_AVS_Friendly %><br/>
			<strong>CVV &#8213;</strong> <%= str_CCV_Friendly %>

			<div class="font-weight-bold mt-3">Billing address:</div>
			<%= rsGetOrder.Fields.Item("billing_name").Value %><br/>
			<%= rsGetOrder.Fields.Item("billing_address").Value %><br/>
			<%= rsGetOrder.Fields.Item("billing_zip").Value %><br/>
			<% end if %>
			
			<% If rsGetOrder.Fields.Item("pay_method").Value = "PayPal" then %>
			<div class="font-weight-bold">PayPal information:</div>
				<strong>E-mail &#8213;</strong> <%= var_paypal_email %><br/>
				
			<% end if %>
			<%= var_message %>
			<div class="mt-3">IP address: <%= rsGetOrder.Fields.Item("IPaddress").Value %></div>
		</div>
		<div class="modal-footer">
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		</div>
	  </div>
	</div>
  </div>



<% else ' unathorized access error %>
	<h5>Not accessible</h5>
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

</body>
</html>
<script type="text/javascript" src="scripts/invoices.js?v=122321"></script>
<% if request.querystring("bo_item") <> "" then %>
	<script type="text/javascript">
		$('.bo_orange_' + <%= request.querystring("bo_item") %>).trigger('click');
	</script>
<% end if %>
<script type="text/javascript">
	<% If var_access_level = "Packaging" OR var_access_level = "Photography" then  %>
			$('.disable-fields select, .disable-fields input, .disable-fields textarea, .disable-fields button').attr("disabled", true);
	<% end if %>
</script>
<%
DataConn.Close()
%>