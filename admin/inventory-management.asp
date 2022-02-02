<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

sortby = request.querystring("sortby")
sortorder = request.querystring("sortorder")

if sortby = "vendor" then
	var_sortby = "vs_wlsl.brand"
elseif sortby = "threshold" then
	var_sortby = "thresh_percent"
elseif sortby = "waiting" then
	var_sortby = "vw_waiting.waiting"
elseif sortby = "out" then
	var_sortby = "out_percent"	
else
	var_sortby = "vs_wlsl.brand"
end if

if sortorder = "asc" then
	var_sortorder = "ASC"
elseif sortorder = "desc" then
	var_sortorder = "DESC"
else
	var_sortorder = "ASC"
end if

' Retrieve current wholesale assets we have on hand
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT  vs_wlsl.brand, vs_wlsl.wholesale_value, vs_wlsl.total_items, vw_out.out, vw_thresh.threshold, vw_waiting.waiting, CAST(vw_thresh.threshold AS float) / CAST(vs_wlsl.total_items AS float) * 100 AS 'thresh_percent', CAST(vw_out.out AS float) / CAST(vs_wlsl.total_items AS float) * 100 AS 'out_percent', vw_open.DateOrdered FROM  vw_vendor_dashboard_wholesale_assets AS vs_wlsl LEFT OUTER JOIN vw_vendor_dashboard_open AS vw_open ON vs_wlsl.brand = vw_open.Brand LEFT OUTER JOIN vw_vendor_dashboard_waiting AS vw_waiting ON vs_wlsl.brand = vw_waiting.brand LEFT OUTER JOIN vw_vendor_dashboard_threshold AS vw_thresh ON vs_wlsl.brand = vw_thresh.brand LEFT OUTER JOIN vw_vendor_dashboard_qtyout AS vw_out ON vs_wlsl.brand = vw_out.brand ORDER BY CASE WHEN vs_wlsl.brand = 'TOTAL' THEN 1 ELSE 2 END, " &  var_sortby & " " & var_sortorder
Set rsGetMainList = objCmd.Execute()
%>
<!DOCTYPE html> 
<html>
<head>
<meta charset="UTF-8">
<title>Inventory management</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">

	<table class="table table-sm table-striped table-hover">
		<thead class="thead-dark">
		<tr>
		<th class="sticky-top" scope="col">Purchase order</th>
		<th class="sticky-top" scope="col">Open</th>
		<th class="sticky-top" scope="col"><a href="?sortby=vendor&amp;sortorder=asc"><i class="fa fa-sort fa-lg mr-2 sort-icon"></i></a>Vendor</th>
		<% 	
		If user_name = "Nathan" or user_name = "Amanda" or user_name = "Ellen" then %>
		<th class="sticky-top" scope="col">
			<div class="small">
				Admin viewing only
			</div>
			Wholesale assets <i class="fa fa-information fa-lg" data-bs-toggle="tooltip" data-bs-placement="top" title="Current wholesale value of all products that are in stock, active, and not custom orders."></i>

	</th>
		<% end if %>
		<th class="sticky-top" scope="col"><a href="?sortby=out&amp;sortorder=desc"><i class="fa fa-sort fa-lg mr-2 sort-icon"></i></a>Out of stock <i class="fa fa-information fa-lg"  data-bs-toggle="tooltip" data-bs-placement="top" title="Amount of items that are currently at 0"></i></th>
		<th class="sticky-top" scope="col"><a href="?sortby=threshold&amp;sortorder=desc"><i class="fa fa-sort fa-lg mr-2 sort-icon"></i></a>Under threshold <i class="fa fa-information fa-lg" data-bs-toggle="tooltip" data-bs-placement="top" title="Amount of items that are equal or less than the threshold amount"></i></th>
		<th class="sticky-top" scope="col"><a href="?sortby=waiting&amp;sortorder=desc"><i class="fa fa-sort fa-lg mr-2 sort-icon"></i></a>Waiting list</th>
		<th class="sticky-top" scope="col">Total items <i class="fa fa-information fa-lg" data-bs-toggle="tooltip" data-bs-placement="top" title="Amount of items that we have listed on the site"></i></th>
		</tr>
		</thead>
<% 
row_id = 1
do while not rsGetMainList.eof 

' Retrieve orders if in the last 3 months
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT  Top(3) Brand, DateOrdered as po_date, PurchaseOrderID FROM TBL_PurchaseOrders WHERE (DateOrdered >= GETDATE() - 90) and Brand = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50, rsGetMainList.Fields.Item("brand").Value))
	Set rs3monthspos = objCmd.Execute()
	
if rsGetMainList.Fields.Item("brand").Value = "TOTAL" then
	apply_class = "font-weight-bold"
else
	apply_class = ""
end if

' only show totals row to ADMINS
If var_access_level = "Admin" and rsGetMainList.Fields.Item("brand").Value = "TOTAL" then 
	admin_class = ""
else
	if rsGetMainList.Fields.Item("brand").Value = "TOTAL" then
		admin_class = "display:none"
	else
		admin_class = ""
	end if
end if
%>
	<tr style="<%= admin_class %>">
		<td>
			<% if rsGetMainList.Fields.Item("brand").Value <> "TOTAL" and rsGetMainList.Fields.Item("brand").Value <> "----" then %>
			<span class="btn btn-sm btn-secondary mr-2 toggle-vendor" id="<%= row_id %>" data-brand="<%= rsGetMainList.Fields.Item("brand").Value %>"><i class="fa fa-angle-down fa-lg vendor-expand<%= row_id %>"></i><i class="fa fa-angle-up fa-lg vendor-expand<%= row_id %>" style="display:none"></i></span><a class="btn btn-sm btn-purple" href="inventory_view.asp?brand=<%= rsGetMainList.Fields.Item("brand").Value %>&amp;resume=yes" target="_blank" title="Create new purchase order" alt="Create new purchase order">New PO</a>
			
			<% if not rs3monthspos.eof then
			i = 1 
			%>
			
			<% while not rs3monthspos.eof 
			
			if rs3monthspos.Fields.Item("po_date").Value >= date() - 30 then
				var_dim = 1
			elseif rs3monthspos.Fields.Item("po_date").Value < date() - 30 and rs3monthspos.Fields.Item("po_date").Value >= date() - 60 then
				var_dim = .8
			else
				var_dim = .6
			end if		
			%>
			<a href="inventory/view_order.asp?ID=<%= rs3monthspos.Fields.Item("PurchaseOrderID").Value %>" target="_blank" title="View purchase order" alt="View purchase order"><span class="badge badge-secondary" style="opacity: <%= var_dim %>"><%= MonthName(Month(rs3monthspos.Fields.Item("po_date").Value),1) %>&nbsp;<%= day(rs3monthspos.Fields.Item("po_date").Value) %></span></a>
			<% 
			rs3monthspos.movenext()
			wend 
			end if %>
			<% end if %>
		</td>
		<td>
		<% if rsGetMainList.Fields.Item("DateOrdered").Value <> "" then %>
		<span class="badge badge-secondary"><%= MonthName(Month(rsGetMainList.Fields.Item("DateOrdered").Value),1) %>&nbsp;<%= day(rsGetMainList.Fields.Item("DateOrdered").Value) %>,&nbsp;<%= year(rsGetMainList.Fields.Item("DateOrdered").Value) %></span>
		<% end if %>
		</td>
		<td class="<%= apply_class %>">
			<a class="btn btn-sm btn-purple mr-2" href="inventory_view.asp?brand=<%= rsGetMainList("brand") %>&amp;readonly=yes" target="_blank" >View stock</a>
			<a href="add_company.asp#<%= rsGetMainList.Fields.Item("brand").Value %>" target="_blank"><%= rsGetMainList.Fields.Item("brand").Value %></a>
		</td>
		<% 	if user_name = "Nathan" or user_name = "Amanda" or user_name = "Ellen" then  %>
		<td class="<%= apply_class %>">
			<%= formatcurrency(rsGetMainList.Fields.Item("wholesale_value").Value,0) %>
		</td>
		<% end if %>
		<td class="<%= apply_class %>">
		<% 	if rsGetMainList.Fields.Item("out").Value <> "" then 
		
		var_out_percent = formatnumber(rsGetMainList.Fields.Item("out_percent").Value,0)
		
		if var_out_percent >= 25 then
			class_out = "badge badge-danger"
		else
			class_out = ""
		end if		
		%>
		<span class="<%= class_out %>"><%= rsGetMainList.Fields.Item("out").Value %> out / <%= var_out_percent %>% out</span>
		<% end if 'if out > 0
		%>
		</td>
		<td class="<%= apply_class %>">
		<% 	if rsGetMainList.Fields.Item("threshold").Value <> "" then
		var_threshold_percent = formatnumber(rsGetMainList.Fields.Item("thresh_percent").Value,0)
		
		if var_threshold_percent >= 25 then
			class_thresh = "badge badge-danger"
		else
			class_thresh = ""
		end if
		%>
		<span class="<%= class_thresh %>"><%= rsGetMainList.Fields.Item("threshold").Value %> under / 
		<%= var_threshold_percent %>% under</span>
		<% end if ' threshold > 0 
		%>
		</td>
		<td class="<%= apply_class %>">
			<% if rsGetMainList.Fields.Item("waiting").Value > 0 then %> <span class="btn btn-sm btn-secondary toggle-waiting" id="<%= row_id %>" data-brand="<%= rsGetMainList.Fields.Item("brand").Value %>"><i class="fa fa-angle-down fa-lg waiting-expand<%= row_id %>"></i><i class="fa fa-angle-up fa-lg waiting-expand<%= row_id %>" style="display:none"></i></span>
			
			<%= rsGetMainList.Fields.Item("waiting").Value %>
			<% end if ' if anyone on waiting list 
			%>
			
			
		</td>
		<td class="<%= apply_class %>">
			<%= formatnumber(rsGetMainList.Fields.Item("total_items").Value,0) %>
			
		</td>
	</tr>
<tbody class="tbody-nohover">
	<tr class="td-expand<%= row_id %>" style="display:none">
		<td colspan="8" class="load<%= row_id %>">
		</td>
	</tr>
	<tr class="tr-waiting-expand<%= row_id %>" style="display:none">
		<td colspan="8" class="waiting-load<%= row_id %>">
		</td>
	</tr>
</tbody>
<% 
row_id = row_id + 1
rsGetMainList.MoveNext()
loop
%> 	
	</table>
</div>
<script type="text/javascript" src="/js/popper.min.js"></script>
<script type="text/javascript">
	$(document).on("click", '.toggle-vendor', function() {
		var brand = $(this).attr("data-brand");
		var row_id = $(this).attr("id");
		$('.vendor-expand' + row_id).toggle();
		$('.td-expand' + row_id).fadeToggle('fast');
	
		$('.load' + row_id).load('/admin/inventory/ajax-vendor-detailed-info.asp?brand=' + encodeURI(brand));
	});
	
	$(document).on("click", '.toggle-waiting', function() {
		var brand = $(this).attr("data-brand");
		var row_id = $(this).attr("id");
		$('.waiting-expand' + row_id).toggle();
		$('.tr-waiting-expand' + row_id).fadeToggle('fast');
	
		$('.waiting-load' + row_id).load('/admin/inventory/ajax-waiting-list-bybrand.asp?brand=' + encodeURI(brand));
	});
</script>
</body>
</html>
