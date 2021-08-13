<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM tbl_closeout_forms WHERE CAST(date_created AS date) >= CAST(GETDATE() - 10 AS date) ORDER BY date_created DESC"
Set rsGetManifests = objCmd.Execute()

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT PackagedBy FROM sent_items WHERE ship_code = N'paid' AND shipped = N'Pending shipment' GROUP BY PackagedBy ORDER BY PackagedBy"
Set rsPrintByPerson = objCmd.Execute()

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT name FROM TBL_AdminUsers WHERE toggle_packer = 1 ORDER BY name ASC"
Set rsGetPackers = objCmd.Execute()
%>
<!DOCTYPE html> 
<html>
<head>
<title>DHL Labels & Closeouts</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
  <div class="container-fluid">
    <div class="row">
      <div class="col-12 col-sm alert alert-secondary m-2">
        <h5>Print Shipping Forms & Close Outs</h5>
        <div class="small">
          <ul>
            <li>To help troubleshoot undefined errors, use <a href="dhl-current-orders.asp" target="_blank">this link</a> to see all current DHL orders. Usually the first 1-4 orders contain the issue. A strange character, missing data, etc.</li>
            <li><span class="text-primary pointer" id="btn-archive-print">Archived printing access</span></li>
          </ul>
        </ul>
<a href="usps/usps-buy-postage.asp" target="_blank">Buy postage (for testing)</a>
        </div>


           
          <div class="alert alert-warning" style="display:none" id="div-archive-print">
            <h5>Archived printing access</h5>
            <a href="print-invoices-dhl.asp" class="btn btn-secondary d-inline-block" target="_blank" role="button">Print ALL DHL invoices (non integrated)</a>
 
              <% if NOT rsPrintByPerson.EOF then
              While NOT rsPrintByPerson.EOF 
                if rsPrintByPerson.Fields.Item("PackagedBy").Value <> "" then 
                %>
                  <a href="dhl/dhl-print-labels.asp?packer=<%= rsPrintByPerson.Fields.Item("PackagedBy").Value %>" class="btn btn-secondary d-inline-block" target="_blank" role="button"><%= rsPrintByPerson.Fields.Item("PackagedBy").Value %> DHL Labels</a>
              <%
              end if
              rsPrintByPerson.MoveNext()
              Wend
              rsPrintByPerson.MoveFirst()
              end if 
              %>
 

          </div>
      </div>
      <div class="col-12 col-sm alert alert-secondary m-2">
        <h5>Assign orders </h5>
        <form id="frm-assign"> 
          <% While NOT rsGetPackers.EOF %>
            <span class="mr-4 d-inline-block"><input class="mr-1" name="Employee" type="checkbox" value="<%= rsGetPackers("name") %>"><%= rsGetPackers("name") %></span>
          <%             
          rsGetPackers.MoveNext()
          Wend %>
          <div class="mt-2">
                  <button class="btn btn-sm btn-primary mt-2 mr-2" id="btn-assign-orders" type="button">Assign orders</button>
                  <span id="msg-assign"></span>
                </div>

        </form>
      </div>
    </div><!-- row -->
    <!-- container -->
        
        <h5 class="rounded p-1 pl-3 mt-2 bg-secondary text-white">Request shipping labels</h5>
              <button type="button" class="btn btn-outline-secondary d-inline-block btn-submit" data-type="requestLabel" data-url="dhl/dhl-request-label-v4.asp?all=yes" >Request DHL Labels<i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="btn-spinner-requestLabel"></i></button>
              <button type="button" id="btn-usps-request-labels" class="btn btn-outline-secondary d-inline-block btn-submit" data-type="requestUSPSLabel" data-url="usps/usps-request-label.asp?all=yes" >Request USPS Labels<i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="btn-spinner-requestUSPSLabel"></i></button>
              <div class="mt-2" id="html-message"></div>
             
        

             <h5 class="rounded p-1 pl-3 mt-4 bg-secondary text-white">Print Shipping Forms</h5>
             <a href="print-invoices-usps.asp?shipper=usps" class="btn btn-outline-secondary d-inline-block" target="_blank" role="button">Print USPS invoices</a>
             <a href="print-invoices-usps.asp?shipper=ups" class="btn btn-outline-secondary d-inline-block" target="_blank" role="button">Print UPS invoices</a>
             <%  if NOT rsPrintByPerson.EOF then
             While NOT rsPrintByPerson.EOF 
              if rsPrintByPerson.Fields.Item("PackagedBy").Value <> "" then
            %>
              <div class="d-inline-block">
              <a href="integrated-shipping-labels.asp?packer=<%= rsPrintByPerson.Fields.Item("PackagedBy").Value %>" class="btn btn-outline-secondary d-inline-block" target="_blank" role="button"><%= rsPrintByPerson.Fields.Item("PackagedBy").Value %></a>
            </div>
            <%
              end if
            rsPrintByPerson.MoveNext()
            Wend
            rsPrintByPerson.MoveFirst()
            end if
            %>
             <h5 class="rounded p-1 pl-3 mt-4 bg-secondary text-white">Manifests / Close Outs</h5>
              <button type="button" class="btn btn-outline-secondary d-inline-block btn-submit" data-type="closeOut" data-url="dhl/dhl-close-out.asp">DHL close day<i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="btn-spinner-closeOut"></i></button>
              <button type="button" class="btn btn-outline-secondary d-inline-block btn-submit" data-type="closeOut-usps" data-url="usps/usps-close-out.asp">USPS Close Out Day<i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="btn-spinner-closeOut-usps"></i></button>


              

<div class="mt-5" id="load-manifests"></div>
<h3>DHL Manifests</h3>
<div id="load-dhl-manifests"></div>
<h3 class="mt-2">USPS Manifests</h3>
<div id="load-usps-manifests"></div>
              
</div>
</body>
</html>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript">
	$('#load-dhl-manifests').load('dhl/dhl-manifest-list.asp');
	$('#load-usps-manifests').load('usps/usps-manifest-list.asp');
	
	$(".btn-submit").click(function(e) {
        var url = $(this).attr("data-url");
        var type = $(this).attr("data-type");
        $('#btn-spinner-' + type).show();
        $('#html-message').html('');

		$.ajax({   
			method: "post",
				url: url,
			dataType: "json",
			statusCode: {
			  408: function() {
				$('#btn-spinner-' + type).hide();
				$('#html-message').append('<div class="alert alert-danger p-2">TIMED OUT || MAX LABELS PROCESSED - Request labels again<br><br>' + json.message + '</div>');
			  }
			}
		})
        .done(function(json, msg) {
			$('#btn-spinner-' + type).hide();
			$('#load-dhl-manifests').load('dhl/dhl-manifest-list.asp');
			$('#load-usps-manifests').load('usps/usps-manifest-list.asp');

			if(json.status == 'success') {
				$('#html-message').append('<div class="alert alert-success p-2">' + json.message + '</div>');
			}
			if(json.status == 'error') {
				$('#html-message').append('<div class="alert alert-danger p-2">' + json.message + '</div>');
			}
            })
            .fail(function(json, msg) {
                $('#btn-spinner-' + type).hide();
                $('#html-message').append('<div class="alert alert-danger p-2">ERROR - Code did not process<br><br>' + json.message + '</div>');
        })
 
    });

        $("#btn-archive-print").click(function() {
          $("#div-archive-print").toggle();
        })

	// Assign orders
  $(document).on("click", "#btn-assign-orders", function(event){
    $('#msg-assign').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');

		$.ajax({
		method: "post",
		url: "packing/assign-orders.asp",
		data: $("#frm-assign").serialize()
		})
		.done(function(msg) {
      $('#msg-assign').html('<span class="alert alert-success p-1">Orders assigned</span>');
		})
		.fail(function(msg) {
      $('#msg-assign').html('<span class="alert alert-danger p-1">Code error assigning orders</span>');
		});
	}); // assign orders

</script>
<%
DataConn.Close()
%>