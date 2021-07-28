<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


if request.form("qty") <> "" then
%>
<!--#include virtual="/checkout/inc_random_code_generator.asp"-->
<!--#include virtual="/includes/inc-dupe-onetime-codes.asp"--> 
<%	

For i = 1 To request.form("qty")

		var_cert_code = getPassword(15, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)
	
		' Call function
		CheckDupe(var_cert_code)
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBLDiscounts (DiscountCode, DiscountDescription, DateExpired, coupon_single_email, DiscountPercent, coupon_single_use, DateAdded, DiscountType, active, dateactive) VALUES (?, 'Admin generated', ?, ?, ?, 1, GETDATE(), 'Percentage', 'A', GETDATE()-1)"
		objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,30,var_cert_code))
		objCmd.Parameters.Append(objCmd.CreateParameter("Expires",200,1,30, request.form("expdate")))
		objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,30,rec_name))
		objCmd.Parameters.Append(objCmd.CreateParameter("Discount",3,1,10, request.form("percentage")))
		objCmd.Execute()

Next

Response.Redirect "one-time-coupons.asp"

end if ' if qty <> ""


Set rsGetCoupons = Server.CreateObject("ADODB.Recordset")
rsGetCoupons.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCoupons.Source = "SELECT * FROM TBLDiscounts WHERE coupon_single_use = 1 AND coupon__single_redeemed = 0 AND coupon_assigned = 0 AND (coupon_single_email = '' OR coupon_single_email IS NULL) ORDER BY DiscountID DESC"
rsGetCoupons.CursorLocation = 3 'adUseClient
rsGetCoupons.LockType = 1 'Read-only records
rsGetCoupons.Open()

%>
<html>
<head>
<title>One time use coupons</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div  class="p-3">
<h4>One time use coupons</h4>

<table class="table table-sm table-hover table-striped">
<thead class="thead-dark">
  <tr>
    <th colspan="5">
	<form class="form-inline">
		<i class="fa fa-plus-circle fa-lg text-success mr-1"></i><span class="text-success mr-5">GENERATE CODES: </span>
		<label for="qty">How many:</label>
		<input class="form-control form-control-sm ml-1 mr-3" name="qty" id=qty" type="number" value="1" min="1" max="100">
		
		<label for="percentage">Discount %</label>
		<input class="form-control form-control-sm ml-1 mr-3"name="percentage" id="percentage" type="number" value="10">

		<label for="expdate">Expires:</label>
		<input class="form-control form-control-sm ml-1 mr-3" name="expdate" id="expdate" type="date">

		<button type="submit" formmethod="post" formaction="one-time-coupons.asp" class="btn btn-sm btn-secondary">Generate</button>
	</form>
	</th>
  </tr>
  <tr>
    <th>Code</th>
    <th>Expires</th>
	<th>Discount %</th>
	<th class="text-center">Assigned?</th>
    <th>Customer e-mail (if any)</th>
  </tr>
</thead>
<tbody class="ajax-update">
<% 
While NOT rsGetCoupons.EOF

if rsGetCoupons.Fields.Item("coupon_assigned").Value = 1 then
	var_checked = "checked"
else
	var_checked = ""
end if

%>
    <tr class="<%= rsGetCoupons.Fields.Item("DiscountID").Value %>">
      <td>
		  <%= rsGetCoupons.Fields.Item("DiscountCode").Value %>
	  <td>
		  <input class="form-control form-control-sm" type="text" value="<%= rsGetCoupons.Fields.Item("DateExpired").Value %>" name="expires_<%= rsGetCoupons.Fields.Item("DiscountID").Value %>" data-column="DateExpired" data-id="<%= rsGetCoupons.Fields.Item("DiscountID").Value %>">
	</td>
	  <td>
		  <input class="form-control form-control-sm" type="text" value="<%= rsGetCoupons.Fields.Item("discountpercent").Value %>" name="percentage_<%= rsGetCoupons.Fields.Item("DiscountID").Value %>" data-column="discountpercent" data-id="<%= rsGetCoupons.Fields.Item("DiscountID").Value %>">
		</td>
      <td class="text-center">
		  <input class="" type="checkbox" value="1" data-unchecked="0" name="assigned_<%= rsGetCoupons.Fields.Item("DiscountID").Value %>" data-column="coupon_assigned" data-id="<%= rsGetCoupons.Fields.Item("DiscountID").Value %>" <%= var_checked %>>
		</td>
	  <td>
		  <input class="form-control form-control-sm" type="text" value="<%= rsGetCoupons.Fields.Item("coupon_single_email").Value %>" name="email_<%= rsGetCoupons.Fields.Item("DiscountID").Value %>" data-column="coupon_single_email" data-id="<%= rsGetCoupons.Fields.Item("DiscountID").Value %>">
		</td>
    </tr>
    <% 
  rsGetCoupons.MoveNext()
Wend
%>
</tbody>
</table>

</div>
</body>
</html>
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript">

	//url to to do auto updating
	var auto_url = "administrative/ajax_update_onetime_coupon.asp"
</script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
	auto_update();
</script>
<%
rsGetCoupons.Close()
Set rsGetCoupons = Nothing
%>
