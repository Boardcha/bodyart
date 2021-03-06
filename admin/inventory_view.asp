<%@LANGUAGE="VBSCRIPT"%>
<% response.Buffer=false
Server.ScriptTimeout=300
 %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

if request.querystring("readonly") = "yes" then
	readonly = "yes"
end if

var_brand = request("brand")

for_how_many_months = 3 'default value
If request("months") <> "" Then 
	for_how_many_months = request("months") 'If there is a user selection overwrite it
Else
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (1) months_to_restock FROM TBL_PurchaseOrders WHERE brand = ? ORDER BY DateOrdered DESC" 
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,100, var_brand))
	Set rsForHowManyMonths = objCmd.Execute()
	If Not rsForHowManyMonths.EOF Then 
		If rsForHowManyMonths("months_to_restock") > 0 Then
			for_how_many_months = rsForHowManyMonths("months_to_restock")
		End If		
	End If
End If	

if var_brand = "Etsy" then
	var_brand = "etsy_stock = 1"
else
	var_brand = "brandname = '" + var_brand + "'"
end if

if readonly <> "yes" then
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
else
	tempid = "0"
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
	objCmd.CommandText = "SELECT inventory.ProductDetailID, jewelry, title, ProductDetail1, ProductDetail2, ProductID, qty, brandname, customorder, wlsl_price, detail_code, type, stock_qty, reorder_amt, picture, largepic, product_active, item_active, DetailPrice, item_order, free, DateLastPurchased, Free_QTY, weight, Length, Gauge, Date_PriceCheck, material, restock_threshold, GaugeOrder, detail_notes, DetailCode, etsy_stock, BinNumber_Detail, date_added, SaleDiscount, ID_BarcodeOrder, ID_Number, ID_SortOrder, DetailID, (restock_threshold - qty) * -1 as thresh_level, (SELECT ISNULL(SUM(ORD.qty), 0) FROM sent_items SNT INNER JOIN TBL_OrderSummary ORD ON SNT.ID = ORD.invoiceID AND ORD.DetailID = inventory.ProductDetailID WHERE SNT.ship_code = 'paid' AND SNT.date_order_placed > DateAdd(month, -" & for_how_many_months & ", GETDATE())) AS sales_from_n_months_back_to_now, (SELECT ISNULL(SUM(ORD.qty), 0) FROM sent_items SNT INNER JOIN TBL_OrderSummary ORD ON SNT.ID = ORD.invoiceID AND ORD.DetailID = inventory.ProductDetailID WHERE SNT.ship_code = 'paid' AND date_order_placed BETWEEN DateAdd(month, -" & for_how_many_months & ", DateLastPurchased) AND DateLastPurchased) as sales_from_n_months_back_to_last_sold_date, sales as sales_from_po_date_received, ISNULL((SELECT po_qty FROM tbl_po_details where po_temp_id = " & tempid & " AND (po_orderid = 0) AND po_detailid = inventory.ProductDetailID), 0) as qty_edited, last_purchase_received, (SELECT TOP(1) po_confirmed FROM tbl_po_details WHERE  (po_detailid = inventory.ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_confirmed ,(SELECT TOP(1) po_manual_adjust FROM tbl_po_details WHERE  (po_detailid = inventory.ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_manual_adjust, (SELECT TOP(1) po_qty_vendor FROM tbl_po_details WHERE (po_detailid = inventory.ProductDetailID) AND (po_orderid = 0) AND (po_temp_id = " + tempid + " )) AS po_qty_vendor, (SELECT TOP(1) po_date_received FROM tbl_po_details WHERE (po_detailid = inventory.ProductDetailID) ORDER BY po_date_received DESC) AS po_date_received, amt_waiting, inventory.location,  inventory.autoclavable, TBL_Barcodes_SortOrder.ID_Description, avg_rating FROM inventory INNER JOIN TBL_Barcodes_SortOrder ON inventory.DetailCode = TBL_Barcodes_SortOrder.ID_Number LEFT OUTER JOIN vw_po_waiting ON inventory.ProductDetailID = vw_po_waiting.DetailID LEFT JOIN TBL_Sales_From_Last_Restock Sales ON Sales.ProductDetailID = inventory.ProductDetailID WHERE " + var_brand + " " + po_filter_title + " " + po_filter_details + " AND customorder <> 'yes' " + po_filter_active + " " + po_filter_status + " ORDER BY " + var_1st_filter + " title ASC, GaugeOrder ASC, ProductID ASC, item_order ASC"
	'Response.Write objCmd.CommandText
	'Response.End 
end if
'Response.Write objCmd.CommandText 
'Response.End
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
<link href="../CSS/fortawesome/css/external-min.css?v=031920" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/popper.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui.min.js"></script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js?v=081223"></script>
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
<% if readonly <> "yes" then %>
<div class="reset">    
	  <button class="btn btn-sm btn-secondary reset_po"type="button">Reset to new</button>
	  
	  <% if request.querystring("resume") <> "" then %>
<% end if %> 
</div>
<% end if '==== if readonly <> "yes" %>


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

<div class="mt-4 mb-2">
	Create Order:
</div>
<form class="ajax-update">
<% if readonly <> "yes" then %>

<div class="form-inline wrapper-createpo">   
		<select class="form-control form-control-sm mr-2" id="for_how_many_months" name="for_how_many_months" style="width:195px">
		<option disabled="disabled" <%If for_how_many_months = 0 Then Response.Write "selected"%>>For how many months:</option>
			<option value="3" <%If for_how_many_months = 3 Then Response.Write "selected"%>>3 months</option>
			<option value="4" <%If for_how_many_months = 4 Then Response.Write "selected"%>>4 months</option>
			<option value="5" <%If for_how_many_months = 5 Then Response.Write "selected"%>>5 months</option>
			<option value="6" <%If for_how_many_months = 6 Then Response.Write "selected"%>>6 months</option>
			<option value="7" <%If for_how_many_months = 7 Then Response.Write "selected"%>>7 months</option>
			<option value="8" <%If for_how_many_months = 8 Then Response.Write "selected"%>>8 months</option>
			<option value="9" <%If for_how_many_months = 9 Then Response.Write "selected"%>>9 months</option>
		</select>
		<textarea class="form-control form-control-sm mr-2" data-column="po_notes" type="text" style="height:31px;min-width:37%;" id="po_notes" name="po_notes" placeholder="Notes:"></textarea><br/>
</div>
<%If for_how_many_months = 0 Then%>
	<div id="months_selection_warning" class="bg-warning text-dark rounded mt-2 p-2" style="width: 896px;">Please select how many months the order is for!</div>
<%End If%>
<div class="form-inline pt-2"> 
		<button class="btn btn-purple mr-4 create_po" type="button">Create order</button>
		<span class="alert-success csv p-1 pl-2 pr-2 rounded" style="display:none">ORDER CREATED</span>
</div>		
<% end if '==== if readonly <> "yes" %>

<div class="text-center paging paging-div">
<!--#include file="inventory/inc-new-po-paging.asp" -->
</div>

<table class="table table-striped table-borderless mt-3">
<thead class="thead-dark">
  <tr class="text-nowrap">
	<th class="sticky-top"></th>
	<% if readonly <> "yes" then %>
	<th class="sticky-top text-right" colspan="2">
		<a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=waiting"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Waiting List
	</th>
	<th class="sticky-top">In pairs</th>
	<th class="sticky-top">Vendor qty</th>
	<th class="sticky-top">Line total</th>
	<% end if %>
	<th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=qty"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>On hand</th>
	<% if readonly <> "yes" then %>
    <th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=max"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Max qty</th>
	<th class="sticky-top"><a href="?brand=<%=Request.QueryString("brand")%>&amp;resume=<%=Request.QueryString("resume")%>&amp;1stfilter=thresh"><i class="fa fa-sort fa-lg sort-icon mr-2"></i></a>Threshold</th>
	<% end if %>
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
row_id = 1
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
	<tr class="bg-secondary">
		<td colspan="6">
		<a href="../productdetails.asp?ProductID=<%=(rsGetDetail.Fields.Item("ProductID").Value)%>" target="_blank"><img src="https://bafthumbs-400.bodyartforms.com/<%=(rsGetDetail.Fields.Item("picture").Value)%>" class="rounded float-left mr-2" style="height:50px;width:50px" <%= var_img_enlarge %>></a>
		
		<a class="text-light h5" href="product-edit.asp?ProductID=<%=(rsGetDetail.Fields.Item("ProductID").Value)%>" target="_blank"><%= rsGetDetail.Fields.Item("title").Value %><% if rsGetDetail.Fields.Item("type").Value <> "None" then %> - <%= rsGetDetail.Fields.Item("type").Value %><% end if %> (<%=(rsGetDetail.Fields.Item("ProductID").Value)%>)</a>

		<% 
		If request.cookies("po-filter-autoclave") = "" Then
			if rsGetDetail("avg_rating") <> "" then 
				var_avg_rating = FormatNumber(rsGetDetail("avg_rating"),1)
				var_avg_percentage = var_avg_rating * 20
			%>
			<span class="rating-box h5 ml-3">
					<span class="rating h5" style="width:<%= var_avg_percentage %>%"></span>
				</span>
			<% end if %>
		<% end if %>

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
			<a href="http://bodyartforms-products.bodyartforms.com/<%= rsGetDetail.Fields.Item("largepic").Value %>" class="btn btn-sm btn-outline-light enlarge" title="<%= rsGetDetail.Fields.Item("material").Value %>">Enlarge image</a>
		</div>
		</td>
		<td>
			<div>
				<label class="font-weight-bold" style="color:white" for="vartype">Status</label>
				<select class="form-control form-control-sm " name="vartype_<%=rsGetDetail("ProductID")%>" data-column="type" data-id="<%=rsGetDetail("ProductID")%>" data-friendly="Status"  style="width:150px">
					<option>None</option>
					<option value="Clearance">Clearance</option>
					<option value="limited">Limited</option>
					<option value="Discontinued">Discontinued</option>
					<option value="One time buy">One time buy</option>
					<option value="Consignment">Consignment</option>
					<option value="<%=(rsGetDetail.Fields.Item("type").Value)%>" selected><%=(rsGetDetail.Fields.Item("type").Value)%></option>
				</select>
			</div>		
		</td>
		<td>
            <div>
                <label class="font-weight-bold" style="color:white" for="discount">Discount</label>
                <select class="form-control form-control-sm discount-select" name="discount_<%=rsGetDetail("ProductID")%>" data-column="SaleDiscount" data-id="<%=rsGetDetail("ProductID")%>" data-friendly="Discount amount" style="width:100px">
                    <option value="<%= (rsGetDetail.Fields.Item("SaleDiscount").Value) %>" selected><% if rsGetDetail.Fields.Item("SaleDiscount").Value = 0 then %>None<%else%><%= (rsGetDetail.Fields.Item("SaleDiscount").Value) %><% end if %></option>
                    <option value="0">None</option>
                    <option value="5">5%</option>
                    <option value="10">10%</option>
                    <option value="15">15%</option>
                    <option value="20">20%</option>
                    <option value="25">25%</option>
                    <option value="30">30%</option>
                    <option value="35">35%</option>
                    <option value="40">40%</option>
                    <option value="45">45%</option>
                    <option value="50">50%</option>
                    <option value="55">55%</option>
                    <option value="60">60%</option>
                    <option value="65">65%</option>
                    <option value="70">70%</option>
                    <option value="75">75%</option>
                    <option value="80">80%</option>
                    <option value="85">85%</option>
                    <option value="90">90%</option>
                    <option value="95">95%</option>
				</select>
			</div>		
		</td>
		<td class="text-center">
			<label class="font-weight-bold" style="color:white" for="autoclavable_<%= i %>">Autoclavable?</label><br>
			<input class="" type="checkbox" value="1" name="autoclavable_<%= i %>" data-id="<%= rsGetDetail.Fields.Item("ProductID").Value %>" data-column="autoclavable" data-friendly="Autoclavable" <%= var_autoclave_checked %>>
		</td>
		<td class="text-center" colspan="5">
			&nbsp;
		</td>
	</tr>
</tbody>
<% end if 
if request.cookies("po-filter-autoclave") = "" then	' hide details if we're autoclave tagging
var_productid = rsGetDetail.Fields.Item("ProductID").Value

'if it's manually adjusted then highlight it 
	if rsGetDetail.Fields.Item("po_manual_adjust").Value = 1 OR rsGetDetail.Fields.Item("po_confirmed").Value = 1 then
		po_manual_adjust = " table-info"
	else
		po_manual_adjust = ""
	end if
%>
<tbody class="<%= rsGetDetail.Fields.Item("type").Value %> <%= var_class_qty %> <%= var_display_active %>" id="tbody_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
  <tr class=" <%= po_manual_adjust %>" id="tr_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
	<td>
	<span class="btn btn-sm btn-secondary mr-2 toggle-product-detail" id="<%= row_id %>" data-detailID="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"><i class="fa fa-angle-down fa-lg product-detail-expand<%= row_id %>"></i><i class="fa fa-angle-up fa-lg product-detail-expand<%= row_id %>" style="display:none"></i></span>
	</td>
	<td class="flex-nowrap" style="min-width:282px!important">
<% 

var_restock = 0		
var_sales = 0	
If rsGetDetail("qty") > 0 Then
	If for_how_many_months <> "" Then
		var_restock = rsGetDetail("sales_from_n_months_back_to_now")
		var_restock = var_restock - rsGetDetail("qty")
		var_sales = rsGetDetail("sales_from_n_months_back_to_now")
	End If
Else
	var_restock = rsGetDetail("sales_from_n_months_back_to_last_sold_date")
	var_sales = rsGetDetail("sales_from_n_months_back_to_last_sold_date")
End If

'If there are customers waiting this item, add it to the puchase quantity
If Not ISNULL(rsGetDetail.Fields.Item("amt_waiting")) Then
	var_restock = var_restock + rsGetDetail("amt_waiting")
End If	

If rsGetDetail("qty_edited") > 0 Then
	var_restock = rsGetDetail("qty_edited")
End If	

If var_restock < 0 Then var_restock = 0

if var_restock > 0 Then show_check = "yes" Else show_check = "no"

if show_check = "no" then
	var_reorder = ""
	if readonly <> "yes" then
%>
	<span class="mr-4">&nbsp;</span>
<% 
	end if
else 
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

	if readonly <> "yes" then
	%>
		<span class="pointer <%= confirmed_check %> mr-2" id="check<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" data-column="po_qty" data-value="<%= var_restock %>" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"><i class="fa fa-check-circle fa-lg"></i></span>
		<% end if' only display if qty hasn't been adjusted 
		%>
		<span class="<%= confirmed_gap %> mr-2" id="checkgap<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"></span>
<% 
	end if '==== if readonly <> "yes"
end if %>	
		<span class="mr-2"><%= rsGetDetail.Fields.Item("ProductDetailID").Value %></span>
		<% if readonly <> "yes" then %>
		<input name="orderqty_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" type="text" class="form-control form-control-sm orderqty d-inline-block" style="width:50px" value="<%= var_restock %>" data-column="po_qty" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"> 
        <span class="mx-1">*</span> 
		<span class="wlsl_price" data-price="<%=FormatNumber((rsGetDetail.Fields.Item("wlsl_price").Value), -1, -2, -0, -2)%>"><%=FormatCurrency((rsGetDetail.Fields.Item("wlsl_price").Value), -1, -2, -0, -2)%></span>  
		<% end if '==== if readonly <> "yes" then %>   
	</td>
	<td style="text-align:center">
		<% If Not IsNULL(rsGetDetail("amt_waiting")) Then
			var_amt_waiting = rsGetDetail("amt_waiting")
		Else
			var_amt_waiting = 0
		End If%>
		<%
		If for_how_many_months =0 Then
			tooltipTitle="Please select for how many months"
		Else	
			If rsGetDetail("qty") < 1 Then strExtension = " <span style='color:yellow'>to last sold date</span>" Else strExtension = ""
			tooltipTitle = "<b>" & var_sales & "</b> sales in last <b>" & for_how_many_months & "</b> months" & strExtension & "<br>" & _
			"On hand: <b>" & rsGetDetail.Fields.Item("qty").Value  & "</b><br>" & _
			"In Waiting List: <b>" & var_amt_waiting  & "</b><br>" & _
			"Last sold date: " & rsGetDetail("DateLastPurchased") 
		End If
		%>
		
		<span data-toggle="tooltip"  data-html="true"
			title="<%=tooltipTitle%>" 
			class="fa fa-information d-inline-block mt-1" style="font-size:22px;vertical-align:middle;"></span>
	<%If rsGetDetail.Fields.Item("amt_waiting").Value > 0 Then %>
		<a class="badge badge-info p-2" href="waitinglist_view.asp?DetailID=<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" target="_blank"><span class="fa fa-user" aria-hidden="true"><sup class="pl-1 font-weight-bold"><%= rsGetDetail.Fields.Item("amt_waiting").Value %></sup></span></a>
	<%End If%>	
	</td>	
	<% if readonly <> "yes" then %>
	<td class="text-center align-middle">
		<input type="checkbox" name="pair_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">	
	</td>
	<td>
	<%
		if rsGetDetail.Fields.Item("po_qty_vendor").Value>0 Then 
			qty_vendor = rsGetDetail.Fields.Item("po_qty_vendor").Value
		Else
			qty_vendor = 0
		End If	
	%>
		<input type="text" name="qty_vendor_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" class="form-control form-control-sm" style="width:50px" value="<%=qty_vendor%>" data-column="po_qty_vendor" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
	</td>
<% ' LEAVE THIS LINE BELOW ALL CONNECTED SO THAT THE JAVASCRIPT WORKS RIGHT TO CREATE THE TOTAL WITH NULL VALUES
%>
<%
if var_restock > 0 then 
	line_total = FormatNumber(var_restock * rsGetDetail.Fields.Item("wlsl_price").Value, -1, -2, -0, -2) 
else 
	line_total = 0
end if 
%>
	<td class="<%= var_reorder %> <%= var_reorder_class %>" id="line_total_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" data-line_total="<%=line_total%>" data-detailid="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>"><% if var_restock > 0 then %><%= line_total%><%else %>&nbsp;<% end if %>
	</td>
	<% end if '===== if readonly <> "yes" %>
    <td style="text-align:center">
		<% if rsGetDetail.Fields.Item("qty").Value <= 0 then
			qty_class = "po_qty0 badge badge-danger font-weight-bold p-2"
		else
			qty_class = "po_qty_above0 badge badge-success font-weight-bold p-2"
		end if %>
		
		<span class="po_qty <%= qty_class %>"><%= rsGetDetail.Fields.Item("qty").Value %></span>
	</td>
	<% if readonly <> "yes" then %>
	<td style="text-align:center">
		<input type="text" name="maxqty_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" class="form-control form-control-sm" style="width:50px" value="<%=(rsGetDetail.Fields.Item("stock_qty").Value)%>" data-column="stock_qty" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
	</td>
	<td style="text-align:center">
		<input type="text" name="threshold_<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>" class="form-control form-control-sm" style="width:50px" value="<%=(rsGetDetail.Fields.Item("restock_threshold").Value)%>" data-column="restock_threshold" data-id="<%= rsGetDetail.Fields.Item("ProductDetailID").Value %>">
	</td>
<% end if %>
	
		<td class="align-middle">	
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
<tbody class="tbody-nohover">
	<tr class="td-expand<%= row_id %> bg-white" style="display:none">
		<td colspan="14" class="load<%= row_id %>">
		</td>
	</tr>
</tbody>	
  <% 
  end if '	if request.querystring("autoclave") <> " yes"
  row_id = row_id + 1
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

<% if readonly <> "yes" then %>
<div class="fixed-bottom h3 m-0 p-2" style="background: rgba(25, 25, 25, .8);color:#ececec">
	TOTAL: $<span id="total"></span>
</div>
<% end if '====if readonly <> "yes" %>
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
	var handlingInProgress = false;
	$(".orderqty").change(function(e){
		if (handlingInProgress) return;
		handlingInProgress=true;
		var id = $(this).attr("data-id");
		var qty_value = $(this).parent().parent().find('.orderqty').val();
		var wlsl_price = $(this).parent().parent().find('.wlsl_price').attr("data-price");
		var new_price = Math.round(qty_value * wlsl_price);
		var chk_pair = $(this).parent().parent().find("[name^='pair']").is(':checked');
		
		if(chk_pair){
			$(this).parent().parent().find("[name^='qty_vendor']").val(Math.round(qty_value/2));
			$(this).parent().parent().find("[name^='qty_vendor']").change();
		}	
		
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
		$(calculateSum);
		handlingInProgress=false;
	});
	
	// Highlight row if qty_vendor field changes
	$("input[name^='qty_vendor']").change(function(e){
		if (handlingInProgress) return;
		handlingInProgress=true;
		var id = $(this).attr("data-id");
		$(this).parent().parent().find(".orderqty").change();
		$(this).closest("tr").removeClass("css_inactive");
		$(this).closest("tr").addClass("table-info");
		$('#reorder_all').hide();
		$('#check' + id).hide();
		$('#checkgap' + id).addClass("empty_check_gap").show();
		// Call calculation function
		$(calculateSum);
		handlingInProgress=false;
	});	

	$("input[name^='pair']").change(function(e){
		if (handlingInProgress) return;
		handlingInProgress=true;
		var id = $(this).attr("data-id");
		var qty_value = $(this).parent().parent().find('.orderqty').val();
		var chk_pair = $(this).is(':checked');
		if(chk_pair){
			$(this).parent().parent().find("[name^='qty_vendor']").val(Math.round(qty_value/2));
			$(this).parent().parent().find("[name^='qty_vendor']").change();
			$('#check' + id).hide();
		}else{
			$(this).parent().parent().find("[name^='qty_vendor']").val("0");
			$(this).parent().parent().find("[name^='qty_vendor']").change();
		}	
		$(this).parent().parent().find('.orderqty').change();
		handlingInProgress=false;
	});
	
	// Clicking checkmark
	$(".disabled_check").click(function(){
		$(this).removeClass("text-secondary");
		$(this).addClass("text-success");
		
		var column_name = $(this).attr("data-column");
		var column_val = $(this).attr("data-value");
		var id = $(this).attr("data-id");
		var tempid = $("#tempid").val();
		$(this).closest("tr").addClass("table-info");
		
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
		if($("#for_how_many_months").val() != null){
			$("#months_selection_warning").hide();
			var brand = $("#brand").val();
			var tempid = $("#tempid").val();
			var pototal = $("#total").html();
			var for_how_many_months = $("#for_how_many_months").val();
			var po_notes = $("#po_notes").val();
			
			$.ajax({
			method: "POST",
			url: "inventory/ajax_update_po_numbers.asp",
			data: {brand: brand, tempid: tempid, pototal: pototal, for_how_many_months: for_how_many_months, po_notes: po_notes}
			})
			$(".csv").show();
		}else {
			alert("Please select for how many months!");
		}
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
	
	function setQueryStringParamValue(key, value){
		var currentUrl = window.location.href;
		var url = new URL(currentUrl);
		url.searchParams.set(key, value); 
		var newUrl = url.href;
		console.log(newUrl);
		return newUrl;
	}

	$("#for_how_many_months").change(function() {
	  window.location.href = setQueryStringParamValue("months", $("#for_how_many_months").val());
	});
	
	$(".discount-select").change(function(){
		var product_id = $(this).attr("data-id")
		$.ajax({
			method: "POST",
			url: "products/ajax_log_product_sales_on_discount_change.asp",
			data: {discount: $(this).val(), product_id: product_id}
		});
	});	
	
});
</script>
<!--#include file="inventory/inc-product-sales-line-graph.inc" -->
<!--#include file="inventory/inc-last-sold-dates-popover.inc" -->
<%
DataConn.Close()
%>

