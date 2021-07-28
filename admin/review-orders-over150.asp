<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../Connections/authnet.asp"-->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM vw_sum_order_orderhistory WHERE over_150 = 1 AND shipped <> 'Pre-order Approved' AND shipped <> 'On Order' AND shipped <> 'ON HOLD' ORDER BY ID DESC"


set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.Open objCmd
rsGetRecords_total = rsGetRecords.RecordCount
%>
<html>
<head>
<title>Review orders over $150</title>
<!--#include file="includes/inc_scripts.asp"-->
<SCRIPT LANGUAGE="JavaScript">


	// Approve hidden pre-approved orders
	$(document).on("click", ".btn-done", function(event){

        invoice_id = $(this).attr('id');

		$.ajax({
		method: "post",
		url: "invoices/ajax-remove-over150-flag.asp",
		data: {invoice_id: invoice_id}
		})
		.done(function(json,msg ) {
            $('#row_' + invoice_id).hide();
		})
		.fail(function(json,msg) {

		});
	});
</script>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="mx-2">
<h5 class="my-3"><%= rsGetRecords_total %> to be reviewed over $150</h5>

<% If Not rsGetRecords.eof Then %>

<table class="table table-striped table-hover small">
<tbody>
<% 
ii = 0
While NOT rsGetRecords.EOF

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS total_orders FROM sent_items WHERE (customer_first = '" + REPLACE(rsGetRecords.Fields.Item("customer_first").Value,"'", "") + "' AND customer_last = '" + REPLACE(rsGetRecords.Fields.Item("customer_last").Value,"'", "") + "') AND ship_code = 'paid'"
Set rsTotalOrders = objcmd.Execute()
%>
	<tr id="row_<%= rsGetRecords.Fields.Item("ID").Value %>">
	  <td>
	  
		  <a href="invoice.asp?ID=<%= rsGetRecords.Fields.Item("ID").Value %>" target="_top" ><%=rsGetRecords.Fields.Item("ID").Value%></a>
          <a class="ml-5" href="order history.asp?var_first=<%=(rsGetRecords.Fields.Item("customer_first").Value)%>&var_last=<%=(rsGetRecords.Fields.Item("customer_last").Value)%>">View history | <%= rsTotalOrders.Fields.Item("total_orders").Value %> paid orders</a>
          <div>Placed on <%= rsGetRecords.Fields.Item("date_order_placed").Value %></div>
          <% If rsGetRecords.Fields.Item("date_sent").Value <> "" then %>
            <div>Shipped on <%= rsGetRecords.Fields.Item("date_sent").Value %></div>
            <div>Assigned to <%= rsGetRecords.Fields.Item("PackagedBy").Value %><% if rsGetRecords.Fields.Item("ScanInvoice_Timestamp").Value <> "" then %><i class="fa fa-package fa-lg ml-2"></i> PACKAGED<% else %> - NOT PACKAGED YET<% end if %></div>
            <% else %>
            Not shipped yet
          <% end if %>

          <div>
            <button class="mt-4 btn btn-sm btn-secondary btn-done" id="<%= rsGetRecords.Fields.Item("ID").Value %>">Reviewed</button>
            </div>
      </td>
      <td>
        <%=(rsGetRecords.Fields.Item("pay_method").Value)%><br>
        <%=(rsGetRecords.Fields.Item("shipping_type").Value)%>&nbsp;&nbsp;&nbsp;<br>
        <%=(rsGetRecords.Fields.Item("UPS_Service").Value)%>
        
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
        <strong>AVS:</strong> <%= str_AVS_Friendly %><br/>
        <strong>CVV:</strong> <%= str_CCV_Friendly %>
        
        <% end if ' only show if paid with a credit card 
        
        %>
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
                <%=(rsGetRecords.Fields.Item("country").Value)%>
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
<b><font size="3"><% if rsGetRecords.Fields.Item("total_less_discounts").Value < 0 then %>0<% else %><%= FormatCurrency(rsGetRecords.Fields.Item("total_less_discounts").Value, -1, -2, -0, -2) %><% end if %></font></b>
<div class="text-primary mt-2"><%=(rsGetRecords.Fields.Item("customer_comments").Value)%></div>
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
        
</form>
</div>
</body>
</html>
<%
rsGetRecords.Close()
%>

