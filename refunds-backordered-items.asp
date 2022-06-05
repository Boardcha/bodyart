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
	decrypted_refund = objCrypt.Decrypt(password, data)
	
	split_refund = split(decrypted_refund, "|")

	invoice_id_hash = split_refund(0)
	var_customer_number = split_refund(1)
	
	Set objCrypt = Nothing
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * from TBL_Refunds_backordered_items WHERE invoice_id = ? AND encrypted_code = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id_hash",3,1,15, invoice_id_hash))
	objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
	set rsCheckRefund = objCmd.Execute()
%>
	<div class="display-5 mb-5">
		Submit for a refund
	</div>
<% if not rsCheckRefund.eof And invoice_id_param = invoice_id_hash then 
	var_refund_id = rsCheckRefund.Fields.Item("id").Value
%>

	<div id="loaded-div">
		<h5 class="mb-1">You have a <%= FormatCurrency(rsCheckRefund.Fields.Item("refund_total").Value) %> refund available</h5>
		<div>Refunds will typically take 5-7 business days to process back to your account.</div>
		<button class="btn btn-primary mt-2" id="btn-process-refund">Click here to process your refund</button><i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="msg-spinner"></i>
		
		<%If var_customer_number = CustID_Cookie AND var_customer_number > 0  Then%>
		<button class="btn btn-secondary mt-2" id="btn-process-store-credit">Click here to issue a store credit</button><i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="msg-spinner"></i>
		<%End If%>
	</div>
	<div id="msg"></div>

<% else %>
	<div class="alert alert-warning">No refund is available to be processed. If you'd like to contact customer service <a class="font-weight-bold" href="/contact.asp">click here</a>.</div>

<% end if ' if a record is found %>

<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">
	$("#btn-process-refund").click(function() {
		$('#btn-process-refund').prop('disabled', true);
		$('#btn-process-store-credit').prop('disabled', true);
		$('#msg-spinner').show();
	
		$.ajax({
		method: "post",
		url: "accounts/ajax-backorder-refunds.asp?encrypted=<%= data %>&id=<%= var_refund_id %>"
		})
		.done(function(msg) {
			$("#msg").addClass("alert alert-success").html("Your refund has been submitted and a confirmation has been sent to the e-mail address on the order.");
			$('#loaded-div').hide();
		})
		.fail(function(msg) {
			$("#msg").addClass("alert alert-danger").html("Error sending form").show();
			$('#btn-process-refund').prop('disabled', false);
			$('#msg-spinner').hide();
		})
	});
	
	$("#btn-process-store-credit").click(function() {
		$('#btn-process-refund').prop('disabled', true);
		$('#btn-process-store-credit').prop('disabled', true);
		$('#msg-spinner').show();
	
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
			$('#msg-spinner').hide();
		})
	});	
	
</script>