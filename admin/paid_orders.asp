<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set DataConn = Server.CreateObject("ADODB.connection")
DataConn.Open MM_bodyartforms_sql_STRING ' CONNECTION STRING FOR ALL PROCEDURES

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = DataConn
rsGetRecords.Source = "SELECT ship_code, shipped, ID, customer_first, customer_last FROM sent_items WHERE ship_code = 'paid' AND (Review_OrderError <> 1 OR  Review_OrderError IS NULL) AND (shipped = 'Pending shipment') ORDER BY ID DESC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()

total_records = rsGetRecords.RecordCount
%>
<html>
<head>
<title>Orders to be shipped</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="px-2">

<h5 class="my-3">
	<%= total_records %> orders to be shipped
</h5>
<button type="button" class="btn btn-primary mb-4" id="btn-set-shipped">Set all orders to shipped <i class="fa fa-spinner fa-2x fa-spin ml-3" id="btn-spinner" style="display:none"></i></button>
<div id="etsy-status"></div>
<div id="orders-status"></div>


</div>
</body>
</html>
<script type="text/javascript" src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript">
	
	
	$("#btn-set-shipped").click(function() {
    $('#btn-spinner').show();

		$.ajax({
		method: "post",
		url: "etsy/etsy-push-tracking.asp"
		})
		.done(function(msg) {
        $("#etsy-status").html('<div class="alert alert-success">Etsy orders have been set to shipped</div>');

          $.ajax({
          method: "post",
          url: "invoices/ajax-set-orders-to-shipped.asp"
          })
          .done(function(msg) {
              $("#orders-status").html('<div class="alert alert-success">All other orders have been set to shipped</div>');
              $('#btn-spinner').hide();
          })
          .fail(function(msg) {
              $("#orders-status").html('<div class="alert alert-danger">Orders were not set to shipped</div>');
              $('#btn-spinner').hide();
          })
		})
		.fail(function(msg) {
        $("#etsy-status").html('<div class="alert alert-danger">Etsy failed</div>');
        $('#btn-spinner').hide();
		})
	});
	
</script>
<%
rsGetRecords.Close()
set rsGetRecords = Nothing
set DataConn = Nothing
%>