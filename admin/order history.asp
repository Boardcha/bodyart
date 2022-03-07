<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

If Request.Querystring("UPS") <> "" then
varUPS = "OR UPS_tracking = '" + Request.Querystring("UPS") + "'"
end if

If Request.Querystring("TransID") <> "" then
varTransID = "OR transactionID = '" + Request.Querystring("TransID") + "'"
end if

If Request.Querystring("custid") <> "" then
varCustid = "OR customer_ID = " + Request.Querystring("custid")
end if
%>
<%
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT sent_items.coupon_code, sent_items.ID, sent_items.customer_first, sent_items.customer_last, sent_items.shipped, sent_items.UPS_tracking, sent_items.USPS_tracking, sent_items.shipping_type, sent_items.ship_code, CAST(sent_items.customer_comments AS NVARCHAR(500)) AS customer_comments, sent_items.date_sent, CAST(sent_items.comments AS NVARCHAR(500)) AS comments, CAST(sent_items.our_notes AS NVARCHAR(500)) AS our_notes, CAST(sent_items.item_description AS NVARCHAR(500)) AS item_description, sent_items.shipping_rate, sent_items.salestax, sent_items.coupon_amt, sent_items.total_preferred_discount, sent_items.total_store_credit, sent_items.total_coupon_discount, sent_items.total_sales_tax, sent_items.total_free_credits, sent_items.total_gift_cert, sent_items.total_returns, sent_items.price, sent_items.transactionID, TBLDiscounts.DiscountPercent FROM sent_items LEFT OUTER JOIN TBLDiscounts ON sent_items.coupon_code = TBLDiscounts.DiscountCode WHERE (customer_first = '" + Request.Querystring("var_first") + "' AND customer_last = '" + Request.Querystring("var_last") + "') OR email = '" + Request.Querystring("var_email") + "' OR ID = '" + Request.Querystring("invoiceno") + "'" + varUPS + " " + varTransID + " " + varCustid + " GROUP BY sent_items.coupon_code, sent_items.ID, sent_items.customer_first, sent_items.customer_last, sent_items.shipped,sent_items.UPS_tracking, sent_items.USPS_tracking, sent_items.shipping_type, sent_items.ship_code, sent_items.date_sent, sent_items.shipping_rate, sent_items.salestax, sent_items.coupon_amt, sent_items.total_preferred_discount, sent_items.total_store_credit, sent_items.total_coupon_discount, sent_items.total_sales_tax, sent_items.total_free_credits, sent_items.total_gift_cert, sent_items.total_returns, sent_items.price, sent_items.transactionID, TBLDiscounts.DiscountPercent, CAST(sent_items.customer_comments AS NVARCHAR(500)), CAST(sent_items.comments AS NVARCHAR(500)), CAST(sent_items.our_notes AS NVARCHAR(500)), CAST(sent_items.item_description AS NVARCHAR(500)) ORDER BY ID DESC"

	set rsGetOrders = Server.CreateObject("ADODB.Recordset")
	rsGetOrders.CursorLocation = 3 'adUseClient
	rsGetOrders.Open objCmd
	rsGetOrders.PageSize = 10 ' not using (possibly needed for pagination)
	intPageCount = rsGetOrders.PageCount ' not using (possibly needed for pagination)


Select Case Request("Action")
	case "<<"
		intpage = 1
	case "<"
		intpage = Request("intpage")-1
		if intpage < 1 then intpage = 1
	case ">"
		intpage = Request("intpage")+1
		if intpage > intPageCount then intpage = IntPageCount
	Case ">>"
		intpage = intPageCount
	case else
		intpage = 1
end select	
	

'rsGetOrders.CursorLocation = 3 'adUseClient
'rsGetOrders.LockType = 1 'Read-only records
'rsGetOrders.Open()
rsGetOrders_numRows = 0
%>
<%
Dim rsGetFlaggedInfo
Dim rsGetFlaggedInfo_cmd
Dim rsGetFlaggedInfo_numRows

Set rsGetFlaggedInfo_cmd = Server.CreateObject ("ADODB.Command")
rsGetFlaggedInfo_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetFlaggedInfo_cmd.CommandText = "SELECT ID, customer_first, customer_last, shipped, ship_code, date_sent FROM sent_items WHERE (shipped = 'FLAGGED' OR shipped = 'CHARGEBACK' or shipped = 'RETURN') AND (customer_first = '" + Request.Querystring("var_first") + "' AND customer_last = '" + Request.Querystring("var_last") + "' OR email = '" + Request.Querystring("var_email") + "' OR ID = '" + Request.Querystring("invoiceno") + "')" 
rsGetFlaggedInfo_cmd.Prepared = true

Set rsGetFlaggedInfo = rsGetFlaggedInfo_cmd.Execute
rsGetFlaggedInfo_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsGetOrders_numRows = rsGetOrders_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsGetOrders_total
Dim rsGetOrders_first
Dim rsGetOrders_last

' set the record count
rsGetOrders_total = rsGetOrders.RecordCount

' set the number of rows displayed on this page
If (rsGetOrders_numRows < 0) Then
  rsGetOrders_numRows = rsGetOrders_total
Elseif (rsGetOrders_numRows = 0) Then
  rsGetOrders_numRows = 1
End If

' set the first and last displayed record
rsGetOrders_first = 1
rsGetOrders_last  = rsGetOrders_first + rsGetOrders_numRows - 1

' if we have the correct record count, check the other stats
If (rsGetOrders_total <> -1) Then
  If (rsGetOrders_first > rsGetOrders_total) Then
    rsGetOrders_first = rsGetOrders_total
  End If
  If (rsGetOrders_last > rsGetOrders_total) Then
    rsGetOrders_last = rsGetOrders_total
  End If
  If (rsGetOrders_numRows > rsGetOrders_total) Then
    rsGetOrders_numRows = rsGetOrders_total
  End If
End If
%>

<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsGetOrders_total = -1) Then

  ' count the total records by iterating through the recordset
  rsGetOrders_total=0
  While (Not rsGetOrders.EOF)
    rsGetOrders_total = rsGetOrders_total + 1
    rsGetOrders.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsGetOrders.CursorType > 0) Then
    rsGetOrders.MoveFirst
  Else
    rsGetOrders.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsGetOrders_numRows < 0 Or rsGetOrders_numRows > rsGetOrders_total) Then
    rsGetOrders_numRows = rsGetOrders_total
  End If

  ' set the first and last displayed record
  rsGetOrders_first = 1
  rsGetOrders_last = rsGetOrders_first + rsGetOrders_numRows - 1
  
  If (rsGetOrders_first > rsGetOrders_total) Then
    rsGetOrders_first = rsGetOrders_total
  End If
  If (rsGetOrders_last > rsGetOrders_total) Then
    rsGetOrders_last = rsGetOrders_total
  End If

End If
%>

<html>
<head>
<title>Customer order history</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>

<!--#include file="admin_header.asp"-->
<div class="px-3">
	<% If Not rsGetFlaggedInfo.EOF Or Not rsGetFlaggedInfo.BOF Then %>
    <h5 class="alert alert-danger mt-3">Customer is flagged for either a return, suspicious order, or chargeback. Check order history for info.</h5>
<% End If ' end Not rsGetFlaggedInfo.EOF Or NOT rsGetFlaggedInfo.BOF %>

<% If rsGetOrders.EOF And rsGetOrders.BOF Then %>
	No orders found
<% End If ' end rsGetOrders.EOF And rsGetOrders.BOF %>

<% If Not rsGetOrders.EOF Then %>
<h4 class="py-4">
   <%=(rsGetOrders_total)%> Orders for <%=(rsGetOrders.Fields.Item("customer_first").Value)%>&nbsp; <%=(rsGetOrders.Fields.Item("customer_last").Value)%>
</h4>
<div class="text-center">
	<!--#include file="invoices/inc_orderhistory_paging.asp" -->
</div>

<% 
	 '======== PAGING
	rsGetOrders.AbsolutePage = intPage
	copy_order_header = ""
	For intRecord = 1 To rsGetOrders.PageSize  
%>
<div class="card bg-light mb-5">
	<h5 class="card-header">
		<a href="invoice.asp?ID=<%=(rsGetOrders.Fields.Item("ID").Value)%>" target="_blank">Invoice #<%=(rsGetOrders.Fields.Item("ID").Value)%></a>

		<span class="small mx-5">
			<%=(rsGetOrders.Fields.Item("shipped").Value)%><% if (rsGetOrders.Fields.Item("ship_code").Value) = "paid" AND (rsGetOrders.Fields.Item("shipped").Value) = "Pending..." then %> SHIPMENT<% end if %>
		</span>
		<span class="small mr-5">
			<%=(rsGetOrders.Fields.Item("date_sent").Value)%>
		</span>
		<span class="small">
			<%=(rsGetOrders.Fields.Item("shipping_type").Value)%>
		</small>
	
	
		<% if (rsGetOrders.Fields.Item("USPS_tracking").Value) <> "" then %>
			<% if instr(rsGetOrders.Fields.Item("shipping_type").Value,"DHL") > 0 then %>
				<span name="<%= rsGetOrders.Fields.Item("USPS_tracking").Value %>" class="usps_tracking btn btn-sm btn-secondary ml-5" data-url="../dhl/dhl-tracking.asp?tracking=">DHL Tracking Details</span>
			<% else %>
				<button class="show_cursor usps_tracking btn btn-sm btn-secondary ml-5" name="<%= rsGetOrders.Fields.Item("USPS_tracking").Value %>" data-url="../usps_tools/usps_tracking.asp?id=">
				Tracking <%=(rsGetOrders.Fields.Item("USPS_tracking").Value)%>
				</button>
			<% end if %>
		<% end if %>
			
			<div class="usps_tracking_info_<%= rsGetOrders.Fields.Item("USPS_tracking").Value %>" style="display:none"></div>
			<% if (rsGetOrders.Fields.Item("UPS_tracking").Value) <> "" then %>
				<span class="small ml-5">UPS tracking # <%=(rsGetOrders.Fields.Item("UPS_tracking").Value)%></span>
			<% end if %>
			</h5>
	<div class="card-body">

		
			<table class="table table-sm small table-hover">
					<% 
					' if notes like gift order, conserve plastic are wanted, then display
					if rsGetOrders.Fields.Item("customer_comments").Value <> "" then %>
						<tr>
							<td>
								
							</td>
							<td colspan="2">
								<span class="bold">Customer comments:</span> <%= rsGetOrders.Fields.Item("customer_comments").Value %>
							</td>
						</tr>
					<% end if %>
					<%
					Dim rsGetOrderDetails
					Dim rsGetOrderDetails_numRows
					
					Set rsGetOrderDetails = Server.CreateObject("ADODB.Recordset")
					With rsGetOrderDetails
					rsGetOrderDetails.ActiveConnection = MM_bodyartforms_sql_STRING
					rsGetOrderDetails.Source = "SELECT OrderDetailID, qty, title, ProductDetail1, ProductID, item_price, PreOrder_Desc, notes, backorder, Gauge, Length, jewelry, returned, anodization_fee FROM QRY_OrderDetails WHERE ID = " & rsGetOrders.Fields.Item("ID").Value & ""
					rsGetOrderDetails.CursorLocation = 3 'adUseClient
					rsGetOrderDetails.LockType = 1 'Read-only records
					rsGetOrderDetails.Open()
					
					LineItem = 0
					SumLineItem = 0
					
					copy_order_details = ""
					copy_line_detail = ""
					copy_totals_line = ""
					copy_totals = ""

					Do While Not.Eof

					if rsGetOrderDetails.Fields.Item("returned").Value = 1 then
							class_returned = " table-danger "
						else
							class_returned = ""
						end if
					%>
						<tr class="<%= class_returned %>">
							<td>
								  <%=(rsGetOrderDetails.Fields.Item("qty").Value)%>
							</td>
							<td>
								  <% if (rsGetOrderDetails.Fields.Item("backorder").Value) = 1 then %><span class="btn btn-sm btn-warning font-weight-bold p-0 px-1 border-0 mr-3">ON BO</span><% end if %><% if (rsGetOrderDetails.Fields.Item("returned").Value) = 1 then %><span class="btn btn-sm btn-danger font-weight-bold p-0 px-1 border-0 mr-3">RETURNED</span><% end if %><a class="text-dark" href="../productdetails.asp?ProductID=<%=(rsGetOrderDetails.Fields.Item("ProductID").Value)%>" target="_blank"><%=(rsGetOrderDetails.Fields.Item("title").Value)%></a>&nbsp;<%=(rsGetOrderDetails.Fields.Item("ProductDetail1").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("PreOrder_Desc").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("notes").Value)%>
							</td>
							<td>
								  $<%= FormatNumber(rsGetOrderDetails.Fields.Item("item_price").Value *rsGetOrderDetails.Fields.Item("qty").Value, -1, -2, -0, -2)%>
								  <% if rsGetOrderDetails("anodization_fee") > 0 then %>
								  + <%= FormatCurrency(rsGetOrderDetails("qty") * rsGetOrderDetails("anodization_fee"), -1, -2, -0, -2) %> color add-on
								  <% end if %>
							</td>
						</tr>
					<%
						copy_line_detail = rsGetOrderDetails.Fields.Item("qty").Value & "&nbsp;&nbsp;|&nbsp;&nbsp;" & rsGetOrderDetails.Fields.Item("title").Value & "&nbsp;" & rsGetOrderDetails.Fields.Item("ProductDetail1").Value & "&nbsp;" & rsGetOrderDetails.Fields.Item("Gauge").Value & "&nbsp;" & rsGetOrderDetails.Fields.Item("Length").Value & "&nbsp;" & rsGetOrderDetails.Fields.Item("PreOrder_Desc").Value & "&nbsp;" & rsGetOrderDetails.Fields.Item("notes").Value &  "&nbsp;&nbsp;&nbsp;&nbsp;$" & FormatNumber(rsGetOrderDetails.Fields.Item("item_price").Value * rsGetOrderDetails.Fields.Item("qty").Value, -1, -2, -0, -2)
						
						copy_order_details = copy_order_details & "&#10;" &  copy_line_detail
					
						LineItem = rsGetOrderDetails.Fields.Item("item_price").Value * rsGetOrderDetails.Fields.Item("qty").Value
						
						SumLineItem = SumLineItem + LineItem
						sum_anodization_fees = sum_anodization_fees + rsGetOrderDetails("qty") * rsGetOrderDetails("anodization_fee")
						.Movenext()
						InvoiceTotal = SumLineItem + (rsGetOrders.Fields.Item("shipping_rate").Value) - (rsGetOrders.Fields.Item("coupon_amt").Value)
					Loop
					End With 
					%>
						<tr>
							<td colspan="2" class="text-right">
								Subtotal
							</td>
							<td>
								<%= FormatCurrency(SumLineItem, -1, -2, -0, -2) %>
							</td>
						</tr>
					
					<%
					copy_totals = "Subtotal: &nbsp;&nbsp;&nbsp;" & FormatCurrency(SumLineItem, -1, -2, -0, -2) & "&#10;" 
					
					rsGetOrderDetails.Close()
					Set rsGetOrderDetails = Nothing
					rsGetOrderDetails_numRows = 0
					
					' Array for invoice totals
					ReDim arrTotals(2,6) 
					
					'arrTotals(col,row)
					arrTotals(0,0) = "10% preferred discount" 
					arrTotals(1,0) = "total_preferred_discount" 
					total_preferred_discount = rsGetOrders.Fields.Item("total_preferred_discount").Value
					arrTotals(2,0) = "&#8722;"
					arrTotals(0,1) = "Coupon discount" 
					arrTotals(1,1) = "total_coupon_discount" 
					total_coupon_discount = rsGetOrders.Fields.Item("total_coupon_discount").Value
					arrTotals(2,1) = "&#8722;" 
					arrTotals(0,2) = "Tax" 
					arrTotals(1,2) = "total_sales_tax" 
					total_sales_tax = rsGetOrders.Fields.Item("total_sales_tax").Value
					arrTotals(2,2) = "&nbsp;&nbsp;"
					arrTotals(0,3) = "Gift certificate" 
					arrTotals(1,3) = "total_gift_cert"
					total_gift_cert = rsGetOrders.Fields.Item("total_gift_cert").Value 
					arrTotals(2,3) = "&#8722;"
					arrTotals(0,4) = "Free gift (USE NOW) credits" 
					arrTotals(1,4) = "total_free_credits" 
					total_free_credits = rsGetOrders.Fields.Item("total_free_credits").Value
					arrTotals(2,4) = "&#8722;"
					arrTotals(0,5) = "Store account credit" 
					arrTotals(1,5) = "total_store_credit"
					total_store_credit = rsGetOrders.Fields.Item("total_store_credit").Value
					arrTotals(2,5) = "&#8722;"
					arrTotals(0,6) = "Order returns" 
					arrTotals(1,6) = "total_returns"
					total_returns = rsGetOrders.Fields.Item("total_returns").Value
					arrTotals(2,6) = "&#8722;"
					
					For i = 0 to UBound(arrTotals, 2) 
					
						if rsGetOrders.Fields.Item(arrTotals(1,i)).Value <> 0 then
					%>
					
						<tr>
							<td colspan="2" class="text-right border-0">
								<%= arrTotals(0,i) %>
							</td>
							<td class="border-0">
								<% 
								var_minus = ""
								if arrTotals(2,i) = "&#8722;" then 
								var_minus = arrTotals(2,i)
								%><%= arrTotals(2,i) %>&nbsp;<% end if %><%= FormatCurrency(rsGetOrders.Fields.Item(arrTotals(1,i)).Value, -1, -2, -0, -2) %>
							</td>
						</tr>
					<% 
					copy_totals_line = arrTotals(0,i) & ":&nbsp;&nbsp;&nbsp;" & var_minus & FormatCurrency(rsGetOrders.Fields.Item(arrTotals(1,i)).Value, -1, -2, -0, -2) & "&#10;"
					copy_totals = copy_totals & "" & copy_totals_line
					
						end if ' if i > 2 or values not 0
					next ' loop through totals array
					
					InvoiceTotal = (SumLineItem + sum_anodization_fees - total_preferred_discount - total_coupon_discount - total_free_credits + rsGetOrders.Fields.Item("shipping_rate").Value + total_sales_tax - total_store_credit - total_gift_cert - total_returns)
					
					%>
						<tr>
							<td colspan="2"  class="text-right border-0">
								Shipping
							</td>
							<td class=" border-0">
								<%= FormatCurrency(rsGetOrders.Fields.Item("shipping_rate").Value, -1, -2, -0, -2) %>
							</td>
						</tr>
						<tr>
							<td colspan="2" class="text-right h5 border-0">
								TOTAL
							</td>
							<td class="order-total h5  border-0">
								<% if InvoiceTotal < 0 then %>0<% else %><%= FormatCurrency(InvoiceTotal, -1, -2, -0, -2) %><% end if %>
							</td>
						</tr>
					</table>
					<%
					copy_totals = copy_totals & "Shipping:&nbsp;&nbsp;&nbsp;" & FormatCurrency(rsGetOrders.Fields.Item("shipping_rate").Value, -1, -2, -0, -2) & "&#10;" & "&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#8212;&#10;TOTAL:&nbsp;&nbsp;&nbsp;" & FormatCurrency(InvoiceTotal, -1, -2, -0, -2) & ""
					copy_order_header = "Invoice # " & rsGetOrders.Fields.Item("ID").Value & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsGetOrders.Fields.Item("shipped").Value & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsGetOrders.Fields.Item("date_sent").Value & "&nbsp;&nbsp;&nbsp;&nbsp;" & rsGetOrders.Fields.Item("shipping_type").Value
					%>
						<button class="btn btn-sm btn-secondary mb-3" id="copy-order" data-clipboard-text="<%= copy_order_header %>&#10;<%= replace(copy_order_details, """", " inch") %>&#10;&#10;<%= copy_totals %>"> <i class="fa fa-content-copy"></i> Copy order</button>

					<%



			' get our private notes for the order
			Set objCmd = Server.CreateObject ("ADODB.Command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "SELECT tbl_invoice_notes.*, TBL_AdminUsers.name FROM TBL_AdminUsers INNER JOIN tbl_invoice_notes ON TBL_AdminUsers.ID = tbl_invoice_notes.user_id WHERE invoice_id = ? ORDER BY date_created DESC" 
			objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,rsGetOrders.Fields.Item("ID").Value))
			Set rs_GetOurNotes = objCmd.Execute()

			if not rs_GetOurNotes.eof then
			%>
				
				<div class="alert alert-success">
				<h5 class="alert-heading">OUR NOTES:</h5>
				<% while not rs_GetOurNotes.eof %>
				<div class="small  mb-3">
				<span class="font-weight-bold mr-5"><%= rs_GetOurNotes.Fields.Item("name").Value %></span><%= rs_GetOurNotes.Fields.Item("date_created").Value %></span><br/>
				<%= rs_GetOurNotes.Fields.Item("note").Value %>
				</div>
				<% rs_GetOurNotes.movenext()
				wend
				%>
				</div>
			<% end if %>
	</div><!-- card body -->
  </div><!-- card wrapper -->



		



  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetOrders.MoveNext()
  
	If rsGetOrders.EOF Then Exit For  ' ====== PAGING
	Next ' ====== PAGING
'Wend
%>
<div class="text-center">
	<!--#include file="invoices/inc_orderhistory_paging.asp" -->
</div>
<% End If ' end Not rsGetOrders.EOF Or NOT rsGetOrders.BOF %>
<script type="text/javascript" src="../js/jquery-2.1.1.min.js"></script>
<script type="text/javascript" src="/js/clipboard.js"></script>
<script>
$(document).ready(function(){

	new Clipboard('#copy-order'); // Clipboard
	
	//	$(".arrow-up").hide();
		
		// Load up tracking information on hover
		$('.usps_tracking').hover(function() {
			var id = $(this).attr('name');
			var url = $(this).attr("data-url");
			$('.usps_tracking_info_' + id).load(url + id);	
		}); 

		// Toggle slide tracking information to display
		$('.usps_tracking').click(function() {
			var id = $(this).attr('name');
			$('.usps_tracking_info_' + id).slideToggle( "slow" );
				$(this).find(".tracking-hide, .tracking-show").toggle();
			}); 
});
</script>
</div><!-- main container div -->
</body>
</html>
<%
rsGetOrders.Close()
%>
<%
rsGetFlaggedInfo.Close()
Set rsGetFlaggedInfo = Nothing
%>
