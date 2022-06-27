<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Refunds"
	page_description = "Process refunds directly for customers"
	page_keywords = ""
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->

<%
	' decrypt refund information
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = request.querystring("hash")
	invoice_id_param = request.querystring("id")
	data = Replace(data, " ", "+") 'Bug fix: IIS converts "+" signs to spaces. We need to convert it back.
	set objCmd = Server.CreateObject("ADODB.Command")
	
	If data <> "" Then
		decrypted_refund = objCrypt.Decrypt(password, data)
		split_refund = split(decrypted_refund, "|")
		invoice_id_hash = split_refund(0)
		ProductDetailID = split_refund(1)
		var_customer_number = split_refund(2)
			
		Set objCrypt = Nothing

		objCmd.ActiveConnection = DataConn
		'objCmd.CommandText = "SELECT * from TBL_Refunds_backordered_items REF WHERE invoice_id = ? AND encrypted_code = ? AND detailID = ?"
		objCmd.CommandText = "SELECT REF.*, (JEW.title + ' ' + ISNULL(DET.Gauge, '') + ' ' + ISNULL(DET.Length, '') + ' ' + ISNULL(DET.ProductDetail1, '')) as description from TBL_Refunds_backordered_items REF LEFT JOIN ProductDetails DET ON REF.ProductDetailID = DET.ProductDetailID LEFT JOIN Jewelry JEW ON JEW.ProductID = DET.ProductID WHERE redeemed = 0 AND REF.invoice_id = ? AND REF.encrypted_code = ? AND REF.ProductDetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id_hash",3,1,15, invoice_id_hash))
		objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
		objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, ProductDetailID))
		set rsCheckRefund = objCmd.Execute()
	
		%>
			<div class="display-5 mb-3">
				Submit for a refund
			</div>
		<% if Not rsCheckRefund.eof And invoice_id_param = invoice_id_hash then 
				var_refund_id = rsCheckRefund.Fields.Item("id").Value
				%>
				<div id="loaded-div">
					<div class="mb-1 font-weight-bold">You have a <%= FormatCurrency(rsCheckRefund.Fields.Item("refund_total").Value) %> refund available for the item below.</div>
					<div class="mb-3">- <%= rsCheckRefund("description") %></div>
					<div id="msg-spinner" style="display:none"><i class="fa fa-spinner fa-spin fa-lg"></i> Processing...</div>
					<div class="refund-buttons">
						<button class="btn btn-primary mt-2" id="btn-process-refund">Click here to process your refund</button>
						<%If var_customer_number = CustID_Cookie AND var_customer_number > 0  Then%>
							<button class="btn btn-secondary mt-2" id="btn-process-store-credit">Click here to issue a store credit</button>
						<%End If%>
					</div>
					<div class="mt-2"><i>- Refunds will typically take 5-7 business days to process back to your account.</i></div>
					<div><i>- Issuing a store credit is processed into your account immediately.</i></div>
				</div>
				<div id="msg" class="mt-2"></div>
		<% else %>
			<% no_refund_found = true %>
		<% end if ' if a record is found %>
	<% else %>
		<% no_refund_found = true %>
	<% end if %>

	<% If no_refund_found Then %>
		<div class="alert alert-warning">No refund is available to be processed. If you'd like to contact customer service <a class="font-weight-bold" href="/contact.asp">click here</a>.</div>
	<% end if %>

	<!--#include virtual="/bootstrap-template/footer.asp" -->
	<script type="text/javascript">
		$("#btn-process-refund").click(function() {
			$('#btn-process-refund').prop('disabled', true);
			$('#btn-process-store-credit').prop('disabled', true);
			$('#msg-spinner').show();
			$('.refun-buttons').hide();
		
			$.ajax({
			method: "post",
			dataType: "json",
			url: "accounts/ajax-backorder-refunds.asp?encrypted=<%= data %>&id=<%= var_refund_id %>"
			})
			.done(function(json, msg) {
				if (json.status == 'success') {
					$("#msg").addClass("alert alert-success").html("Your refund has been submitted and a confirmation has been sent to the e-mail address on the order.");
					$('#loaded-div').hide();
				}else{
					$("#msg").addClass("alert alert-danger").html("The transaction was unsuccessful. " + json.error + " Please contact customer service at help@bodyartforms.com or call us at (877) 223-5005").show();
					$('#btn-process-refund').prop('disabled', false);
					$('#btn-process-store-credit').prop('disabled', false);
					$('#msg-spinner').hide();			
				}
			})
			.fail(function(json, msg) {
				$("#msg").addClass("alert alert-danger").html("The transaction was unsuccessful. Please contact customer service at help@bodyartforms.com or call us at (877) 223-5005").show();
				$('#btn-process-refund').prop('disabled', false);
				$('#btn-process-store-credit').prop('disabled', false);
				$('#msg-spinner').hide();
			})
		});
		
		$("#btn-process-store-credit").click(function() {
			$('#btn-process-refund').prop('disabled', true);
			$('#btn-process-store-credit').prop('disabled', true);
			$('#msg-spinner').show();
			$('.refun-buttons').hide();
		
			$.ajax({
			method: "post",
			url: "accounts/ajax-backorder-store-credit.asp?encrypted=<%= data %>&id=<%= var_refund_id %>"
			})
			.done(function(msg) {
				$("#msg").addClass("alert alert-success").html("Your store credit has been issued and a confirmation has been sent to the e-mail address on the order.");
				$('#loaded-div').hide();
			})
			.fail(function(msg) {
				$("#msg").addClass("alert alert-danger").html("Error sending form").show();
				$('#btn-process-refund').prop('disabled', false);
				$('#btn-process-store-credit').prop('disabled', false);
				$('#msg-spinner').hide();
			})
		});	
	</script>