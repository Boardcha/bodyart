<%@LANGUAGE="VBSCRIPT"%>
<% response.Buffer=false
Server.ScriptTimeout=300
 %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


var_brand = request("brand")

if var_brand = "Etsy" then
	var_brand = "etsy_stock = 1"
else
	var_brand = "brandname = '" + var_brand + "'"
end if
	
' Create a new purchase order on page load if it's not resuming the last one
if request.querystring("resume") = "yes" and request.cookies(var_brand) <> "" then
	' Continue to use what was originally assigned to cookie below if it's not a new order
	tempid = request.cookies(var_brand)
else
	' Insert a new temp PO id # to use while paging through and creating order since it will need to be saved into database
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_po_temp_ids DEFAULT VALUES"
	objCmd.Execute()
	
	' Retrieve the newest temp PO # to use for saving order details
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT  TOP (1) po_temp_id FROM tbl_po_temp_ids ORDER BY po_temp_id DESC" 
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,100, var_brand))
	Set rsGetTempPONum = objCmd.Execute()

	' Set cookie for brand to newest temp #
	response.cookies(var_brand) = rsGetTempPONum.Fields.Item("po_temp_id").Value
	tempid = request.cookies(var_brand)
end if

' If we need to reset the order refresh the page after the new temp ID has been created
if request.querystring("reset") = "yes" then
	Response.Redirect "?resume=yes&brand=" & request.querystring("brand")
end if

if request.cookies("po-filter-status") = "All stock" then
	po_filter_status = ""
elseif request.cookies("po-filter-status") = "" or request.cookies("po-filter-status") = "Regular stock only" then
	po_filter_status = "AND (type <> 'limited' and type <> 'One time buy' and type <> 'Clearance' and type <> 'Discontinued')"
elseif request.cookies("po-filter-status") = "Limited stock only" then
	po_filter_status = "AND (type LIKE '%limited%')"
elseif request.cookies("po-filter-status") = "Discontinued stock only" then
	po_filter_status = "AND (type = 'Discontinued')"
elseif request.cookies("po-filter-status") = "Clearance stock only" then
	po_filter_status = "AND (type = 'Clearance')"
elseif request.cookies("po-filter-status") = "One time buys only" then
	po_filter_status = "AND (type = 'one time buy')"
elseif request.cookies("po-filter-status") = "Not sold in last 6 months" then
	po_filter_status = "AND DateLastPurchased < DATEADD(month, -6, GETDATE())"
end if


if request.cookies("po-filter-active") = "Showing inactives" then
	po_filter_active = ""
elseif request.cookies("po-filter-active") = "" or request.cookies("po-filter-active") = "Hiding inactives" then
	po_filter_active = "AND item_active = 1 AND product_active = 1"
end if

' Order by sort column selected
if request.querystring("1stfilter") = "qty" then
	var_1st_filter = "qty ASC,"
elseif request.querystring("1stfilter") = "thresh" then
	var_1st_filter = "thresh_level ASC,"
elseif request.querystring("1stfilter") = "max" then
	var_1st_filter = "stock_qty ASC,"
elseif request.querystring("1stfilter") = "lastbought" then
	var_1st_filter = "DateLastPurchased DESC,"
	var_1st_filter = "stock_qty ASC,"
elseif request.querystring("1stfilter") = "waiting" then
	var_1st_filter = "vw_po_waiting.amt_waiting DESC,"
else
	var_1st_filter = ""
end if

' Filter by keywords / title 
if request.querystring("keywords_title") <> "" then
	keyword_array = split(request.querystring("keywords_title")," ")
	For Each strItem In keyword_array
		sql_build_title = sql_build_title & " AND title LIKE '%" & strItem & "%'"
	next
	'sql_build_title = Mid(sql_build_title,6)	'Strip first AND from string
	po_filter_title = sql_build_title
	'response.write  po_filter_title
end if

' Filter by keywords / details 
if request.querystring("keywords_details") <> "" then
	keyword_array = split(request.querystring("keywords_details")," ")
	For Each strItem In keyword_array
		sql_build_detail = sql_build_detail & " AND ProductDetail1 LIKE '%" & strItem & "%'"
	next
	'sql_build_detail = Mid(sql_build_detail,6)	'Strip first AND from string
	po_filter_details = sql_build_detail
	'response.write  po_filter_details
end if



Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
if request.cookies("po-filter-autoclave") = "yes" then
	' if we need to mass tag autoclave items then don't pull in all the details
	objCmd.CommandText = "SELECT * FROM jewelry WHERE " + var_brand + " AND autoclavable = 1 ORDER BY title ASC"
elseif request.cookies("po-filter-autoclave") = "tag" then
	' if we need to mass tag autoclave items then don't pull in all the details
	objCmd.CommandText = "SELECT * FROM jewelry WHERE " + var_brand + " " + po_filter_title + " " + po_filter_details + " AND autoclavable = 0 ORDER BY title ASC"
else
	objCmd.CommandText = "SELECT *, (restock_threshold - qty) * -1 as thresh_level,(SELECT TOP(1) po_confirmed FROM tbl_po_details WHERE  (po_detailid = ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_confirmed ,(SELECT TOP(1) po_manual_adjust FROM tbl_po_details WHERE  (po_detailid = ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_manual_adjust,(SELECT TOP(1) po_qty FROM tbl_po_details WHERE (po_detailid = ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_qty, amt_waiting, inventory.location,  inventory.autoclavable, TBL_Barcodes_SortOrder.ID_Description FROM inventory INNER JOIN TBL_Barcodes_SortOrder ON inventory.DetailCode = TBL_Barcodes_SortOrder.ID_Number LEFT OUTER JOIN vw_po_waiting ON inventory.ProductDetailID = vw_po_waiting.DetailID WHERE " + var_brand + " " + po_filter_title + " " + po_filter_details + " AND customorder <> 'yes' " + po_filter_active + " " + po_filter_status + " " + po_filter_qty + " ORDER BY " + var_1st_filter + " title ASC, GaugeOrder ASC, ProductID ASC, item_order ASC"
end if



'objCmd.CommandText = "SELECT *, (restock_threshold - qty) * -1 as thresh_level,(SELECT TOP(1) po_confirmed FROM tbl_po_details WHERE  (po_detailid = ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_confirmed ,(SELECT TOP(1) po_manual_adjust FROM tbl_po_details WHERE  (po_detailid = ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_manual_adjust,(SELECT TOP(1) po_qty FROM tbl_po_details WHERE (po_detailid = ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_qty, amt_waiting, inventory.location,  inventory.autoclavable, TBL_Barcodes_SortOrder.ID_Description FROM inventory INNER JOIN TBL_Barcodes_SortOrder ON inventory.DetailCode = TBL_Barcodes_SortOrder.ID_Number LEFT OUTER JOIN vw_po_waiting ON inventory.ProductDetailID = vw_po_waiting.DetailID WHERE " + var_brand + " AND customorder <> 'yes' " + po_filter_active + " " + po_filter_status + " " + po_autoclave_tag + " " + po_filter_qty + " ORDER BY " + var_1st_filter + " title ASC, GaugeOrder ASC, ProductID ASC, item_order ASC"
set rsGetDetail = Server.CreateObject("ADODB.Recordset")
rsGetDetail.CursorLocation = 3 'adUseClient
rsGetDetail.Open objCmd
rsGetDetail.PageSize = 50 ' not using (possibly needed for pagination)
intPageCount = rsGetDetail.PageCount ' not using (possibly needed for pagination)
total_records = rsGetDetail.RecordCount

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

Set rsGetCompanyInfo_cmd = Server.CreateObject ("ADODB.Command")
rsGetCompanyInfo_cmd.ActiveConnection = DataConn
rsGetCompanyInfo_cmd.CommandText = "SELECT website FROM TBL_Companies WHERE name = ?" 
rsGetCompanyInfo_cmd.Prepared = true
rsGetCompanyInfo_cmd.Parameters.Append rsGetCompanyInfo_cmd.CreateParameter("param1", 200, 1, 100, var_brand) ' adVarChar

Set rsGetCompanyInfo = rsGetCompanyInfo_cmd.Execute
%>
<html>
<head>
<title><%= request.querystring("brand") %> : View stock</title>
<link rel="stylesheet" href="../js-fancybox2/source/jquery.fancybox.css?v=2.1.5" type="text/css" media="screen" />
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript" src="/js/popper.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui.min.js"></script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">

	//url to to do auto updating
	var auto_url = "inventory/ajax_update_inventory_view.asp"
</script>
</head>
<body>
<!--#include file="admin_header.asp"-->
<script type="text/javascript" src="/js-fancybox2/source/jquery.fancybox.pack.js?v=2.1.5"></script>
<div class="p-3 create-po">
<input type="hidden" id="tempid" value="<%= tempid %>">
<input type="hidden" id="brand" value="<%= request("brand") %>" />
<h5>
<% if NOT rsGetCompanyInfo.eof then %>
<a href="<%= rsGetCompanyInfo.Fields.Item("website").Value %>" target="_blank">
<% end if %>
<%= Request("brand") %></a> (<%=total_records%> total items)
    <% if Request.Querystring("brand") = "Diablo consignment" OR Request.Querystring("brand") = "Oracle consignment" OR Request.Querystring("brand") = "Royal Organic Jewelry" then %>
	&nbsp;&nbsp;&nbsp;
	<a href="inventory_consignment.asp?brand=<%= Request.Querystring("brand") %>" class="button_small_grey button_shrink">Write a check</a>
    <% end if %>
</h5>
<div class="reset">    
	  <button class="btn btn-sm btn-secondary reset_po"type="button">Reset to new</button>
	  
	  <% if request.querystring("resume") <> "" then %>
<% end if %> 
</div>


<div class="mt-4 mb-2">
	Current filters: <%= request.cookies("po-filter-status") %>&nbsp;&nbsp;&nbsp;<%= request.cookies("po-filter-active") %>&nbsp;&nbsp;&nbsp;<%= request.cookies("po-filter-qty") %>
</div>
<form class="form-inline" id="frm_filters" method="get" action="#">
	<select class="form-control form-control-sm mr-2" id="filter_status">
	<option disabled="disabled" selected="selected">Filter by status:</option>
		<option value="All stock">Show all</option>
		<option value="Regular stock only">Regular stock</option>
		<option value="Limited stock only">Limited</option>
		<option value="Clearance stock only">Clearance</option>
		<option value="Discontinued stock only">Discontinued</option>
		<option value="One time buys only">One offs</option>
		<option value="Not sold in last 6 months">Not sold in last 6 months</option>
	</select>
	<select class="form-control form-control-sm mr-2" id="filter_active">
		<option disabled="disabled" selected="selected">Inactive / Active: </option>
		<option value="Hiding inactives">Hide inactives</option>
		<option value="Showing inactives">Show inactives</option>
	</select>
	<select class="form-control form-control-sm mr-2" id="filter_autoclave">
		<option disabled="disabled" selected="selected">Autoclave tagging: </option>
		<option value="">Show all</option>
		<option value="tag">Not tagged</option>
		<option value="yes">Tagged as autoclave</option>
	</select>
	<input class="form-control form-control-sm mr-2" type="text" name="keywords_title" value="<%= request.querystring("keywords_title") %>" placeholder="Title keyword(s)">
	<input class="form-control form-control-sm mr-2" type="text" name="keywords_details" value="<%= request.querystring("keywords_details") %>" placeholder="Detail keyword(s)">
	<button class="btn btn-sm btn-secondary" type="submit">Search keyword(s)</button>

	<input type="hidden" name="brand" value="<%= request.querystring("brand") %>">
	<input type="hidden" name="resume" value="<%= request.querystring("resume") %>">
	<input type="hidden" name="autoclave" value="<%= request.querystring("autoclave") %>">
	<span class="tmp-loader"></span>
</form>

<form class="ajax-update">
<div class="wrapper-createpo">      
	  <button class="btn btn-purple mr-4 create_po" type="button">Create order</button>
	<span class="alert-success csv" style="display:none">
ORDER CREATED
	</span>
</div>

<div class="text-center paging paging-div">
<!--#include file="inventory/inc-new-po-paging.asp" -->
</div>
<div class="loader-div" style="display:none"></div>
<table class="table table-striped table-borderless mt-3">
<thead class="thead-dark">
  <tr class="text-nowrap">
    <th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Re-order</th>
	<th class="sticky-top">Line total</th>
	<th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=qty"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>On hand</th>
    <th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=max"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Max qty</th>
	<th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=thresh"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Threshold</th>
	<th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=waiting"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Waiting</th>
    <th class="sticky-top">Item information</th>
	<th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=lastbought"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Last sold</th>
	<th class="sticky-top">Vendor SKU</th>
	<th class="sticky-top">Item notes</th>
	<th class="sticky-top">Active</th>
  </tr>
</thead>
<% 
var_productid = "not set yet"
i = 0
if not rsGetDetail.eof then
rsGetDetail.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetDetail.PageSize

if request.cookies("po-filter-autoclave") = "" then	' hide details if we're autoclave tagging
if rsGetDetail.Fields.Item("qty").Value <= 1 then
	var_class_qty = " qty_filter_1"
else
	var_class_qty = " qty_filter_all"
end if

end if '	autoclave tagging

if request.cookies("po-filter-autoclave") <> "" then	' enlarge photo for autoclave tagging
	var_img_enlarge = "style=""width:125px;height:125px"""
end if
%>
<%
if var_productid <> rsGetDetail.Fields.Item("ProductID").Value then
%>
<tbody class="tbody_header <%= rsGetDetail.Fields.Item("type").Value %> <%= var_class_qty %>">
	<tr>
		<td colspan="11" class="bg-secondary">
		<a href="../productdetails.asp?ProductID=<%=(rsGetDetail.Fields.Item("ProductID").Value)%>" target="_blank"><img src="https://bafthumbs-400.bodyartforms.com/<%=(rsGetDetail.Fields.Item("picture").Value)%>" class="rounded float-left mr-2" style="height:50px;width:50px" <%= var_img_enlarge %>></a>
		
		<a class="text-light h5" href="product-edit.asp?ProductID=<%=(rsGetDetail.Fields.Item("ProductID").Value)%>" target="_blank"><%= rsGetDetail.Fields.Item("title").Value %><% if rsGetDetail.Fields.Item("type").Value <> "None" then %> - <%= rsGetDetail.Fields.Item("type").Value %><% end if %> (<%=(rsGetDetail.Fields.Item("ProductID").Value)%>)</a>
			<% 	if rsGetDetail.Fields.Item("autoclavable").Value = 1 then
					var_autoclave_checked = "checked"
				else
					var_autoclave_checked = ""
				end if

			%>		
		<div class="pt-1 tag-autoclave">
		<%	
		if request.cookies("po-filter-autoclave") = "" then
			i = rsGetDetail.Fields.Item("ProductDetailID").Value
		else
			i = i + 1
		end if
		%>
		<input type="checkbox" value="1" name="autoclavable_<%= i %>" data-id="<%= rsGetDetail.Fields.Item("ProductID").Value %>" data-column="autoclavable" data-friendly="Autoclavable" <%= var_autoclave_checked %>> Autoclavable? <a href="http://bodyartforms-products.bodyartforms.com/<%= rsGetDetail.Fields.Item("largepic").Value %>" class="ml-3 enlarge" title="<%= rsGetDetail.Fields.Item("material").Value %>">img</a>
		</div>
		</td>
	</tr>
</tbody>
<% end if 
if request.cookies("po-filter-autoclave") = "" then	' hide details if we're autoclave tagging
var_productid = rsGetDetail.Fields.Item("ProductID").Value

'if it's manually adjusted then highlight it 
	if rsGetDetail.Fields.Item("po_manual_adjust").Value = 1 and rsGetDetail.Fields.Item("po_confirmed").Value = 0 then
		po_manual_adjust = " table-info"
	else
		po_manual_adjust = ""
	end if
%>
<tbody class="<%= rsGetDetail.Fields.Item("type").Value %> <%= var_class_qty %> <%= var_display_active %>" id="tbody_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
  <tr class=" <%= po_manual_adjust %>" id="tr_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
	<td class="form-inline flex-nowrap">
<% 
show_check = "no"
if rsGetDetail.Fields.Item("restock_threshold").Value <> 0 then
	if rsGetDetail.Fields.Item("restock_threshold").Value >= rsGetDetail.Fields.Item("qty").Value then
		var_restock = rsGetDetail.Fields.Item("stock_qty").Value - rsGetDetail.Fields.Item("qty").Value
		show_check = "yes"
	else
		var_restock = 0
	end if
else
	if rsGetDetail.Fields.Item("stock_qty").Value - rsGetDetail.Fields.Item("qty").Value > 0 then
		var_restock = rsGetDetail.Fields.Item("stock_qty").Value - rsGetDetail.Fields.Item("qty").Value
		show_check = "yes"
	else
		var_restock = 0
	end if
end if

' overwrite calculated qty restock amount if we've already saved our own value in the database
	if rsGetDetail.Fields.Item("po_qty").Value <> 0 then
		var_restock = rsGetDetail.Fields.Item("po_qty").Value
	end if

if show_check = "no" then
	var_reorder = ""
	
%>
	<span class="mr-4">&nbsp;</span>
<% else 
	var_reorder = "reorder_all"
	
	'if it's been confirmed then show green check
	if rsGetDetail.Fields.Item("po_confirmed").Value = 1 then
		confirmed_check = "text-success enabled_check"
	else 
		confirmed_check = "text-secondary disabled_check"
	end if

	'if it's been manually adjusted enable check
	if rsGetDetail.Fields.Item("po_manual_adjust").Value = 1 then
		confirmed_check = "text-success"
	end if	
	
if rsGetDetail.Fields.Item("po_manual_adjust").Value = 0 and rsGetDetail.Fields.Item("po_confirmed").Value = 0 then
	confirmed_gap = "no-display"
else
	confirmed_gap = ""
	%>
		<span class="pointer <%= confirmed_check %> mr-2" id="check<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" data-column="po_qty" data-value="<%= var_restock %>" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"><i class="fa fa-check-circle fa-lg"></i></span>
		<% end if' only display if qty hasn't been adjusted 
		%>
		<span class="<%= confirmed_gap %> mr-2" id="checkgap<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"></span>
<% 
end if %>	
		<span class="mr-2"><%= rsGetDetail.Fields.Item("ProductDetailID").Value %></span>
		<input name="orderqty_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" type="text" class="form-control form-control-sm orderqty" style="width:50px" value="<%= var_restock %>" data-column="po_qty" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"> 
        <span class="mx-1">*</span> 
		<span class="wlsl_price" data-price="<%=FormatNumber((rsGetDetail.Fields.Item("wlsl_price").Value), -1, -2, -0, -2)%>"><%=FormatCurrency((rsGetDetail.Fields.Item("wlsl_price").Value), -1, -2, -0, -2)%></span>     
	</td>
<% ' LEAVE THIS LINE BELOW ALL CONNECTED SO THAT THE JAVASCRIPT WORKS RIGHT TO CREATE THE TOTAL WITH NULL VALUES
%>
	<td class="<%= var_reorder %> <%= var_reorder_class %>" id="line_total_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" data-line_total="<% if var_restock > 0 then %><%= FormatNumber(var_restock * rsGetDetail.Fields.Item("wlsl_price").Value, -1, -2, -0, -2) %><% else %>0<% end if %>" data-detailid="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"><% if var_restock > 0 then %><%= FormatCurrency(var_restock * rsGetDetail.Fields.Item("wlsl_price").Value, -1, -2, -0, -2) %><%else %>&nbsp;<% end if %>
	</td>
    <td style="text-align:center">
		<% if rsGetDetail.Fields.Item("qty").Value <= 0 then
			qty_class = "po_qty0 badge badge-danger font-weight-bold p-2"
		else
			qty_class = "po_qty_above0 badge badge-success font-weight-bold p-2"
		end if %>
		
		<span class="po_qty <%= qty_class %>"><%= rsGetDetail.Fields.Item("qty").Value %></span>
	</td>
	<td style="text-align:center">
		<input type="text" name="maxqty_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" class="form-control form-control-sm" style="width:50px" value="<%=(rsGetDetail.Fields.Item("stock_qty").Value)%>" data-column="stock_qty" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
	</td>
	<td style="text-align:center">
		<input type="text" name="threshold_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" class="form-control form-control-sm" style="width:50px" value="<%=(rsGetDetail.Fields.Item("restock_threshold").Value)%>" data-column="restock_threshold" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
	</td>

   
	<td style="text-align:center">
		<a class="badge badge-info p-2" href="waitinglist_view.asp?DetailID=<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" target="_blank"><%= rsGetDetail.Fields.Item("amt_waiting").Value %></a>

	</td>	
		
		<td>	
			<span class="badge badge-secondary p-1" <% 
			if rsGetDetail.Fields.Item("item_active").Value = 1 then
			%>style="display:none"<% end if %> id="active_status_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">INACTIVE</span>
			<% if (rsGetDetail.Fields.Item("type").Value) <> "None" then %><%=(rsGetDetail.Fields.Item("type").Value)%>&nbsp;&nbsp;&nbsp;<% end if %><%= rsGetDetail.Fields.Item("gauge").Value %>&nbsp;<%= rsGetDetail.Fields.Item("length").Value %>&nbsp;<%=(rsGetDetail.Fields.Item("ProductDetail1").Value)%>
		</td>
		<td>
			<% if rsGetDetail.Fields.Item("DateLastPurchased").Value <> "" then %>
				<span role="button" class="date_expand" id="last_sold_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" data-container="body" data-toggle="popover" data-placement="left" data-html="true" data-trigger="focus" data-content='Loading <i class="fa fa-spinner fa-spin ml-3"></i>' data-detailid="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
					<%= FormatDateTime(rsGetDetail.Fields.Item("DateLastPurchased").Value,2)%>
				</span>
			<% end if %>
		</td>
		<td>
			<input type="text" name="sku_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" class="form-control form-control-sm" value="<%=(rsGetDetail.Fields.Item("detail_code").Value)%>" data-column="detail_code" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<input type="text" name="notes_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" class="form-control form-control-sm" value="<% if rsGetDetail.Fields.Item("detail_notes").Value <> "" then %><%= Server.HTMLEncode(rsGetDetail.Fields.Item("detail_notes").Value) %><% end if %>" data-column="detail_notes" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<% 	if (rsGetDetail.Fields.Item("item_active").Value) = 1 then
					var_checked = "checked"
				else
					var_checked = ""
				end if

			%>
			<input name="active_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" type="checkbox" value="1" <%= var_checked %>  data-column="active" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" data-unchecked="0" data-friendly="Detail active" alt="<%= rsGetDetail.Fields.Item("ID_Description").Value %>&nbsp;&nbsp;<%= rsGetDetail.Fields.Item("location").Value %>" title="<%= rsGetDetail.Fields.Item("ID_Description").Value %>&nbsp;&nbsp;<%= rsGetDetail.Fields.Item("location").Value %>">
		</td>
	</tr>
</tbody>
  <% 
  end if '	if request.querystring("autoclave") <> " yes"
  rsGetDetail.MoveNext()
If rsGetDetail.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING

end if ' if not rsGetDetail.eof
%>
</table>
</form>
<div class="text-center paging paging-div mb-5">
<!--#include file="inventory/inc-new-po-paging.asp" -->
</div>
</div>

<div class="fixed-bottom h3 m-0 p-2" style="background: rgba(25, 25, 25, .8);color:#ececec">
	TOTAL: $<span id="total"></span>
</div>
</body>
</html>
<script type="text/javascript">
$(document).ready(function(){

	$("a.enlarge").fancybox({
		type : 'image',
		nextEffect: 'fade',
		prevEffect: 'fade',
		padding: 0, // remove white border
	});

	auto_update(); // run function to update fields when tabbing out of them
	
	// Sum to get grand total
	function calculateSum() {
		$.ajax({
		method: "post",
		dataType: "json",
		url: "inventory/ajax-total-purchase-order.asp",
		data: {tempid: $("#tempid").val()}
		})
		.done(function( json, msg ) {
			// Write total to div at bottom of page
			$('#total').html(json.po_total);
		})
		.fail(function(json, msg) {
			console.log("fail");
		});	
	};
	
	calculateSum(); // have this run on page load
	
	// Reset the purchase order to a new one
	$(".reset_po").click(function(){
		Cookies.set($("#brand").val(), null, { path: '/' });
		$('.tmp-loader').load('/admin/inventory/ajax-reset-po-delete-items.asp?tempid=' + encodeURI($("#tempid").val()));
		window.location.href = "?reset=yes&brand=" + encodeURI($("#brand").val());
	}); // End reset po

	
	// Initialize bootstrap popovers
	$(function () {
		$('[data-toggle="popover"]').popover()
	  })

	 // Close popover if clicking outside of it
	  $('body').on('click', function () {
		$('.popover').popover('hide');
	});

	// START last sold date expand and load ----------------------------------------
    $(".date_expand").click(function(){	
			
        var detailid = $(this).attr("data-detailid");
       $('.popover').popover('hide');
     //    console.log(detailid);	


		$('.loader-div').load("products/ajax_last_sold_dates.asp", {detailid: detailid}, function() {
            $("#last_sold_" + detailid).attr('data-content', $('.loader-div').html());
            $("#last_sold_" + detailid).popover('show');
        });	 
	}); // END last sold date expand and load ----------------------------------------
	
	// Highlight row if qty field changes
	$(".orderqty").change(function(){
		var id = $(this).attr("data-id");
		var qty_value = $(this).parent().find('.orderqty').val();
		var wlsl_price = $(this).parent().find('.wlsl_price').attr("data-price");
		var new_price = Math.round(qty_value * wlsl_price);
		
		// Write new caculations into the data- attribute
		$('#line_total_' + id).html('$' + new_price);
		$('#line_total_' + id).attr('data-line_total', new_price);
		$('#line_total_' + id).addClass('confirmed');
	
		$(this).closest("tr").removeClass("css_inactive");
		$(this).closest("tr").addClass("table-info");
		$('#reorder_all').hide();
		$('#check' + id).hide();
		$('#checkgap' + id).addClass("empty_check_gap").show();
		// Call calculation function
		$(calculateSum)
	});

	// Clicking checkmark
	$(".disabled_check").click(function(){		
		$(this).removeClass("text-secondary");
		$(this).addClass("text-success");
		
		var column_name = $(this).attr("data-column");
		var column_val = $(this).attr("data-value");
		var id = $(this).attr("data-id");
		var tempid = $("#tempid").val();

		$.ajax({
		method: "POST",
		url: auto_url,
		data: {id: id, column: column_name, value: column_val, tempid: tempid, confirmed:"yes"}
		})
		.done(function( msg ) {
			$('#line_total_' + id).addClass('confirmed');
			calculateSum();
		//	console.log("Success");
		//	console.log( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id);
		})
		.fail(function(msg) {
			alert("The re-order amount did not save. Try changing the re-order quantity manually.");
		//	console.log( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
		});
		
	});
	
	// Auto submit filter form after changing select menus
	$('#frm_filters select').change(function() {
        var filter_name = $(this).attr('id');
		var filter_value = $(this).val(); 
		$('.tmp-loader').load('/admin/inventory/ajax-po-set-filters.asp?' + filter_name + '=' + encodeURI(filter_value), function() {
			location.reload();
		});
	});
	

	// After pressing create order button, load in ajax file and show download link
	$('.create_po').click(function() {
		var brand = $("#brand").val();
		var tempid = $("#tempid").val();
		var pototal = $("#total").html();
		
		$.ajax({
		method: "POST",
		url: "inventory/ajax_update_po_numbers.asp",
		data: {brand: brand, tempid: tempid, pototal: pototal}
		})
		$(".csv").show();
	});	
	
		// Toggle grey row change for active/inactive
	$("input[name^=active_]").change(function(){
	
		var id = $(this).attr("data-id");

			if ($(this).prop("checked")) { // Get values if it's a checkbox
			
					$('#active_status_' + id).hide();
				} else {
				
					$('#active_status_' + id).show();
				}	
	}); // Toggle grey row change for active/inactive
	

		
});
</script>
<%
DataConn.Close()
%>
