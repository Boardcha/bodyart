<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Refunds"
	page_description = "Process refunds directly for customers"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<%
	' decrypt refund information
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = request.querystring("id")
	decrypted_refund = objCrypt.Decrypt(password, data)
	
	split_refund = split(decrypted_refund, "|")

	invoice_id = split_refund(0)
	refund_total = split_refund(1)

	Set objCrypt = Nothing

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * from tbl_redeemable_refunds WHERE invoice_id = ? AND refund_total = ? AND encrypted_code = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, invoice_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("refund_total",6,1,20, refund_total))
	objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
	set rsCheckRefund = objCmd.Execute()

%>


<div class="display-5 mb-5">
	Submit for a refund
</div>
<% if not rsCheckRefund.eof then 
var_refund_id = rsCheckRefund.Fields.Item("id").Value
%>

<div id="loaded-div">
	<h5 class="mb-1">You have a <%= FormatCurrency(rsCheckRefund.Fields.Item("refund_total").Value) %> refund available</h5>
	<div>Refunds will typically take 5-7 business days to process back to your account.</div>
	<button class="btn btn-primary mt-2" id="btn-process-refund">Click here to process your refund</button><i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="msg-spinner"></i>
</div>
<div id="msg"></div>

<% else %>
<div class="alert alert-warning">No refund is available to be processed. If you'd like to contact customer service <a class="font-weight-bold" href="/contact.asp">click here</a>.
</div>

<% end if ' if a record is found %>
<!--#include virtual="/bootstrap-template/footer.asp" -->

<script type="text/javascript">
	
	
	$("#btn-process-refund").click(function() {
		$('#btn-process-refund').prop('disabled', true);
		$('#msg-spinner').show();
	
		$.ajax({
		method: "post",
		url: "accounts/ajax-order-refunds.asp?encrypted=<%= data %>&id=<%= var_refund_id %>"
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
	
</script>