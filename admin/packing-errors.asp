<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

if request.querystring("packer") = "" then
  var_packer = "none"
else
  var_packer = request.querystring("packer") 
end if

'------------- GET PACKER NAMES -----------------
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT name FROM TBL_AdminUsers WHERE toggle_packer = 1 ORDER BY name ASC"
Set rsGetPackers = objCmd.Execute()
%>
<!--#include file="packing/inc-error-formula.asp" -->
<!DOCTYPE html>
<html>
<head>
  <% if request.cookies("admindarkmode") <> "on" then %> 
  <link href="/CSS/baf.min.css?v=040220" id="lightmode" rel="stylesheet" type="text/css" />
<% else %>
  <link href="/CSS/baf-dark.min.css?v=050820" id="darkmode" rel="stylesheet" type="text/css" />
<% end if %>
<script src="https://use.fortawesome.com/dc98f184.js"></script>
  <meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<title>Packing error report</title>
</head>
<!--#include file="admin_header.asp"-->
<body>
  <div class="p-3">
<h5><%= request.querystring("packer") %> packing error report</h5>
<FORM class="my-4" action="#" method="get" name="frm-search">
    <div class="form-inline mb-2">
        From:
        <input class="form-control form-control-sm ml-1" name="date1" type="date" value="<%= var_date1 %>">
        
        <span class="ml-4 mr-1">To:</span>
        <INPUT class="form-control form-control-sm" name="date2" type="date" value="<%= var_date2 %>">
        </p>
        <input class="btn btn-sm btn-purple ml-4" type="submit" name="Submit" value="Search">
    </div>
<% if var_access_level = "Admin" OR var_access_level = "Manager" OR user_name = "Andres" then 
 %>
<div class="alert alert-warning w-50">
    <div class="font-weight-bold">Admin viewing only - <%= total_orders %> total orders</div>
    <select class="form-control form-control-sm" name="packer">
      <% if request.querystring("packer") <> "" then %>
      <option value="<%= request.querystring("packer") %>" selected><%= request.querystring("packer") %></option>
      <% end if %>
      <% While NOT rsGetPackers.EOF %>
        <option value="<%= rsGetPackers("name") %>"><%= rsGetPackers("name") %></option>
      <%             
    rsGetPackers.MoveNext()
    Wend %>
    </select>
  </div>
<% end if %>
</FORM>

<div class="alert alert-info">
  <h4><%= total_errors %> errors<span class="mx-4">|</span><%= var_error_percentage %></h4>
  <span class="font-weight-bold mr-1">Flip-flop:</span><%= Error_flip_total %>
  <span class="font-weight-bold ml-3 mr-1">Mis-matched:</span><%= Error_matching_total %>
  <span class="font-weight-bold ml-3 mr-1">Broken:</span><%= Error_broken_total %>
  <span class="font-weight-bold ml-3 mr-1">Missing:</span><%= Error_missing_total %>
  <span class="font-weight-bold ml-3 mr-1">Wrong:</span><%= Error_wrong_total %>
  <span class="font-weight-bold ml-3 mr-1">Misc:</span><%= Error_misc_total %>
</div>

<table class="table table-striped table-borderless table-hover">
<%
While NOT rsGetErrors.EOF 
%>
    <tr id="row-<%= rsGetErrors.Fields.Item("OrderDetailID").Value %>">
      <td><% if var_access_level = "Admin" OR var_access_level = "Manager" OR user_name = "Andres" then 
        %>
        <i class="pointer text-danger mr-4 fa fa-trash-alt remove-error" data-orderdetailid="<%= rsGetErrors.Fields.Item("OrderDetailID").Value %>"></i>
        <% end if %>
        <a href="../../admin/invoice.asp?ID=<%= rsGetErrors.Fields.Item("ID").Value %>" target="_blank"><b><%=(rsGetErrors.Fields.Item("ID").Value)%></b></a><br>

    <strong><%=(rsGetErrors.Fields.Item("item_problem").Value)%>&nbsp;
<% If (rsGetErrors.Fields.Item("item_problem").Value) = "Missing" then %>
 <%=(rsGetErrors.Fields.Item("ErrorQtyMissing").Value)%>&nbsp;&nbsp;Scanned:&nbsp;<%=(rsGetErrors.Fields.Item("TimesScanned").Value)%>
 <% end if %><br>
            Shipped: <%= FormatDateTime(rsGetErrors.Fields.Item("date_sent").Value,1)%>&nbsp;&nbsp;&nbsp;Placed: <%=FormatDateTime((rsGetErrors.Fields.Item("date_order_placed").Value),2)%></strong><br>
        <%=(rsGetErrors.Fields.Item("ErrorDescription").Value)%>
<br>
        <%=(rsGetErrors.Fields.Item("qty").Value)%> | <%=(rsGetErrors.Fields.Item("title").Value)%>&nbsp; <%=(rsGetErrors.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetErrors.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetErrors.Fields.Item("ProductDetail1").Value)%> &nbsp;<%=(rsGetErrors.Fields.Item("notes").Value)%></td>
    </tr>

<% 
  rsGetErrors.MoveNext()
Wend
%>
</table>
<%
rsGetErrors.Close()
Set rsGetErrors = Nothing
%>
</div>
</body>
</html>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript">
$(".remove-error").click(function() {
    var orderdetailid = $(this).attr('data-orderdetailid');
    console.log(orderdetailid);
		$.ajax({
      method: "post",
      url: "packing/remove-error.asp",
      data: {orderdetailid: orderdetailid}
		})
		.done(function(msg) {
			$('#row-' + orderdetailid).fadeOut('slow');
		})
		.fail(function(msg) {
			alert("CODE ERROR");
		})
	});
	
</script>

