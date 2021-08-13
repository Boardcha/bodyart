<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% response.Buffer = false %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../functions/iif.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

	if request.querystring("add-new-product") = "yes" then
	
		set CopyProduct = Server.CreateObject("ADODB.Command")
		CopyProduct.ActiveConnection = DataConn
		CopyProduct.CommandText = "INSERT INTO jewelry(picture, largepic, active, new_page_date, date_added, added_by) VALUES ('nopic.gif', 'nopic.gif', " & 0 & ", '" & now() & "', '" & now() & "', '" & user_name & "')"
		CopyProduct.Execute() 
		
		Set objCmd = Server.CreateObject ("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TOP 1 ProductID FROM jewelry ORDER BY ProductID DESC" 
		Set rsGetID = objCmd.Execute()
		
		response.redirect "product-edit.asp?ProductID=" & rsGetID.Fields.Item("ProductID").Value
		
	end if ' add a new empty product

if Request.QueryString("ProductID") <> "" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM jewelry WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,Request.QueryString("ProductID")))
	Set rs_getproduct = objCmd.Execute()
	
	noproduct = ""
	if rs_getproduct.BOF and rs_getproduct.EOF then
		page_title = "Product not found"
		noproduct = "Product not found"
	else
		page_title = rs_getproduct.Fields.Item("title").Value
	end if
	
	if request("filter_detailid") <> "" then
		filter_detailid = "AND ProductDetailID = " & request("filter_detailid") & ""
	else
		filter_detailid = ""
	end if
	
	if request("filter_active") = "active" then
		filter_active = "AND active = 1"
		filter_select_text = "Showing active only"
	elseif request("filter_active") = "inactive" then
		filter_active = "AND active = 0"
		filter_select_text = "Showing inactive only"
	else
		filter_active = ""
		filter_select_text = "Showing active &amp; inactive"
	end if

	if request("filter_gauge") = "20g" then
		filter_gauge = "AND (Gauge = '20g' OR Gauge = '18g')"
	elseif request("filter_gauge") = "16g" then
		filter_gauge = "AND Gauge = '16g'"
	elseif request("filter_gauge") = "14g" then
		filter_gauge = "AND (Gauge = '14g' OR Gauge = '14g/12g')"
	elseif request("filter_gauge") = "12g" then
		filter_gauge = "AND Gauge = '12g'"
	elseif request("filter_gauge") = "10g" then
		filter_gauge = "AND Gauge = '10g'"
	elseif request("filter_gauge") = "8g" then
		filter_gauge = "AND Gauge = '8g'"
	elseif request("filter_gauge") = "6g" then
		filter_gauge = "AND Gauge = '6g'"
	elseif request("filter_gauge") = "4g" then
		filter_gauge = "AND Gauge = '4g'"		
	elseif request("filter_gauge") = "2g" then
		filter_gauge = "AND Gauge = '2g'"		
	elseif request("filter_gauge") = "0g" then
		filter_gauge = "AND Gauge = '0g'"		
	elseif request("filter_gauge") = "00g" then
		filter_gauge = "AND (Gauge = '00g' OR Gauge = '00g/9mm' OR Gauge = '00g/9.5mm' OR Gauge = '00g/10mm')"					
	elseif request("filter_gauge") = "7/16""" then
		filter_gauge = "AND Gauge = '7/16""'"		
	elseif request("filter_gauge") = "1/2""" then
		filter_gauge = "AND Gauge = '1/2""'"
	elseif request("filter_gauge") = "9/16""" then
		filter_gauge = "AND Gauge = '9/16""'"		
	elseif request("filter_gauge") = "5/8""" then
		filter_gauge = "AND Gauge = '5/8""'"
	elseif request("filter_gauge") = "3/4""" then
		filter_gauge = "AND Gauge = '3/4""'"	
	elseif request("filter_gauge") = "7/8""" then
		filter_gauge = "AND Gauge = '7/8""'"	
	elseif request("filter_gauge") = "1""" then
		filter_gauge = "AND Gauge = '1""'"	
		
	elseif request("filter_gauge") = "Between 1 inch - 2 inch" then
		filter_gauge = "AND (Gauge LIKE '%1-%')"
	elseif request("filter_gauge") = "2 inch - 3 inch" then
		filter_gauge = "AND (Gauge LIKE '%2-%' OR Gauge = '2""' OR Gauge = '3""')"

	elseif request("filter_gauge") = "odd_small" then
		filter_gauge = "AND (Gauge = '13g' OR Gauge = '11g' OR Gauge = '9g' OR Gauge = '7g' OR Gauge = '5g' OR Gauge = '3g' OR Gauge = '1g')"
	elseif request("filter_gauge") = "odd_large" then
		filter_gauge = "AND (Gauge = '11/16""' OR Gauge = '13/16""' OR Gauge = '15/16""')"
	elseif request("filter_gauge") = "odd mm above 00g" then
		filter_gauge = "AND (Gauge LIKE '%mm%' AND Gauge <> '00g/9mm' AND Gauge <> '00g/9.5mm' AND Gauge <> '00g/10mm')"		
		
		
		
	else
		filter_gauge = ""
	end if
	
	' Get ALL details and also total count for details
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM ProductDetails INNER JOIN TBL_GaugeOrder ON COALESCE (ProductDetails.Gauge, '') = COALESCE (TBL_GaugeOrder.GaugeShow, '') INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number LEFT OUTER JOIN tbl_images ON ProductDetails.img_id = tbl_images.img_id WHERE ProductID = ? " & filter_active & " " & filter_gauge & " " & filter_detailid & " ORDER BY active DESC, item_order ASC, GaugeOrder ASC, price ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,Request.QueryString("ProductID")))
	
	set rs_getdetails = Server.CreateObject("ADODB.Recordset")
	rs_getdetails.CursorLocation = 3 'adUseClient
	rs_getdetails.Open objCmd
	rs_getdetails.PageSize = 50 ' not using (possibly needed for pagination)
	intPageCount = rs_getdetails.PageCount ' not using (possibly needed for pagination)

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

	'Get total count of active details
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM ProductDetails INNER JOIN TBL_GaugeOrder ON COALESCE (ProductDetails.Gauge, '') = COALESCE (TBL_GaugeOrder.GaugeShow, '') INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ProductID = ? AND active = 1 " & filter_gauge & "  ORDER BY active DESC, item_order ASC, GaugeOrder ASC, price ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,Request.QueryString("ProductID")))
	
	set rs_getActivedetails = Server.CreateObject("ADODB.Recordset")
	rs_getActivedetails.CursorLocation = 3 'adUseClient
	rs_getActivedetails.Open objCmd
	
	
	' Page count variables ---------------------------
	var_total_details = rs_getdetails.RecordCount
	var_total_active_details = rs_getActivedetails.RecordCount
	var_total_inactive_details = var_total_details - var_total_active_details
	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT companyID, name FROM TBL_Companies WHERE display_AddEdit = 'yes' AND type = 'jewelry' ORDER BY name ASC"
	Set rs_getbrand = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT tag FROM TBL_Product_Tags ORDER BY tag ASC"
	Set rs_getTags = objCmd.Execute()	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT category_name, category_tag FROM TBL_Categories ORDER BY category_name ASC"
	Set rs_getCategories = objCmd.Execute()		

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT material_name FROM TBL_Materials ORDER BY material_name ASC"
	Set rs_getMaterials = objCmd.Execute()	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT material_name FROM TBL_Materials WHERE toggle_wearable=1 ORDER BY material_name ASC"
	Set rs_getWearableMaterials = objCmd.Execute()	

	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID, GaugeShow FROM TBL_GaugeOrder WHERE ID <> 91 ORDER BY GaugeOrder ASC" 
	Set rsGetGauges = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT country FROM TBL_Countries WHERE origin_toggle = 1 ORDER BY country ASC"
	Set rs_getOriginCountries = objCmd.Execute()


	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_Barcodes_SortOrder" 
	Set rs_getsections = objCmd.Execute()
	
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP(40) * FROM tbl_images WHERE product_id = ? ORDER BY img_id ASC" 
	objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,Request.QueryString("ProductID")))
	Set rs_product_images = objCmd.Execute()

end if 'if Request.QueryString("ProductID") <> ""

if rs_getproduct.Fields.Item("new_page_date").Value >= now()-45 then
	new_active = "btn-primary"
	new_text = "Remove from new"
else
	new_active = "btn-secondary"
	new_text = "Add to new"
end if

if rs_getproduct.Fields.Item("active").Value = 1 then
    var_active_class = "alert-success"
else
    var_active_class = "alert-danger"
end if


'Create array for drop down menus
color_array = array("amber", "aqua", "black", "blue", "bone", "brass", "bronze", "brown", "clear", "copper", "dark-blue", "dark-purple", "fuchsia", "hider", "image", "iridescent", "gold", "glow", "gray", "green", "lavender", "light-blue", "lime", "magenta", "metallic", "navy", "neon", "opalescent", "orange", "pattern", "pink", "purple", "rainbow", "red", "rose-gold", "silver", "skin-tone", "tan", "teal", "translucent", "turquoise", "white", "yellow")
%>
					
<!DOCTYPE html> 
<html>
<head>
<meta charset="UTF-8">
<link rel="stylesheet" type="text/css" href="css/chosen.min.css" />
<link rel="stylesheet" type="text/css" href="/css/redactor.css" />
<link rel="stylesheet" type="text/css" href="css/dropzone.css" />
<title>Edit <%= page_title %></title>
<% if request.querystring("tagging") = "yes" then %>
<script type="text/javascript">
$(document).ready(function(){
	// Expand all automatically if tagging
		$(".show-less").toggleClass("details-border");
		$(".expanded-details").toggle('slide');
		var oldText = $(this).text();
		var newText = $(this).data('text');
		$(this).text(newText).data('text',oldText);
}); // end document ready
</script>
<% end if %>
</head>
<body>
<!--#include file="admin_header.asp"-->
<style>
	/* Change bootstrap checkbox from blue color to grey */
	.custom-checkbox .custom-control-input:checked ~ .custom-control-label::before {
    	background-color: #6c757d;
		border-color: #6c757d;
	}

	.chosen-choices{padding-top:0!important;padding-bottom:0!important}
	.css-product-edit .enlarge_footer_image img {border: solid 3px black; width: 250px; max-height: 300px; verflow: hidden;}
	.css-product-edit .thumb-delete:hover {border: 1px solid #FF0000;}
	.css-product-edit .border-toggle-on {border-bottom:solid 4px #6E6E6E!important}
	#edit_images_link{display: none;}
	#update_main_img, #update_400w_img{position: absolute; float: right; right:16px; top:15px;}	
</style>

<% if noproduct <> "" or loggedin = "" then 'if product is found AND user is logged in display information %>
<%= noproduct %><%= logged_message %>
<% else %>
    <div class="container mt-3 css-product-edit ajax-update" style="max-width:100%">
    <div class="row">

    <div class="col-sm pr-4 small">       
            <div class="container p-0 mb-2">
                <div class="row">
					<div class="col-auto">
						<a href="../productdetails.asp?ProductID=<%=(rs_getproduct.Fields.Item("ProductID").Value)%>" target="_blank"><img id="main_img" src="http://bodyartforms-products.bodyartforms.com/<%=(rs_getproduct.Fields.Item("picture").Value)%>" width="90" height="90"> </a>
					  </div>
                    <div class="col"><h4>#<%=(rs_getproduct.Fields.Item("ProductID").Value)%>
                        <a class="btn btn-sm btn-secondary d-inline-block" href="?ProductID=<%= rs_getproduct.Fields.Item("ProductID").Value - 1%>"><i class="fa fa-angle-left fa-lg"></i></a>
						<a class="btn btn-sm btn-secondary d-inline-block" href="?ProductID=<%= rs_getproduct.Fields.Item("ProductID").Value + 1%>"><i class="fa fa-angle-right fa-lg"></i></a></h4> 
						<div class="">Added <%=(rs_getproduct.Fields.Item("date_added").Value)%> by <%= rs_getproduct("added_by") %></div>
						<div>
							<% if rs_getproduct("reviewed_by_1") <> "" then %>
								<div>
								Reviewed by <%= rs_getproduct("reviewed_by_1") %><span class="ml-2"><%= rs_getproduct("review_date_1") %></span>
								</div>
							<% end if %>
							<% if rs_getproduct("reviewed_by_2") <> "" then %>
								<div>
								Reviewed by <%= rs_getproduct("reviewed_by_2") %><span class="ml-2"><%= rs_getproduct("review_date_2") %></span>
								</div>
							<% end if %>
							<% if ISNULL(rs_getproduct("reviewed_by_1")) OR ISNULL(rs_getproduct("reviewed_by_2"))  then %>
							<button class="btn btn-sm btn-info py-0 mb-1" id="reviewed">Reviewed</button>
							<span id="reviewed-msg"></span>
							<% end if %>
						</div>
					</div>
                </div>
				</div>				

				<button class="btn btn-sm <%= new_active %>" id="new-toggle" type="button" data-id="<%=(rs_getproduct.Fields.Item("ProductID").Value)%>"><%= new_text %></button>

				<button id="duplicate" class="btn btn-sm btn-secondary" type="button">
				Duplicate</button>

				<button class="btn btn-sm btn-secondary" type="button" id="show_combine" data-toggle="modal" data-target="#modal-combine">Combine</button>

				<div class="mt-3" style="display:none" id="duplicate-show-buttons">
					<button id="duplicate-product" class="btn btn-sm btn-outline-secondary d-inline-block" type="button" data-id="<%=(rs_getproduct.Fields.Item("ProductID").Value)%>">Product only</button>
					<button id="duplicate-all"  class="btn btn-sm btn-outline-secondary d-inline-block" type="button" data-id="<%=(rs_getproduct.Fields.Item("ProductID").Value)%>">Product + details</button>
				</div>

				<input class="form-control form-control-sm" name="product-id" id="productid" type="hidden" value="<%= rs_getproduct.Fields.Item("ProductID").Value %>">
				
				<div class="form-group mt-3">
					<select name="active" class="form-control form-control-sm d-inline-block <%= var_active_class %>" data-column="active" data-friendly="Active">
						<option selected value="<%=(rs_getproduct.Fields.Item("active").Value)%>" ><% if (rs_getproduct.Fields.Item("active").Value) = 1 then %>Active<% else%>Inactive<% end if %></option>
						<option value="1">Active</option>
						<option value="0">Inactive</option>
					</select>
				</div>

            <div class="form-group">
                <label class="font-weight-bold" for="title">Product title:</label>
                <input class="form-control form-control-sm" name="title" type="text" id="title" value="<%=(rs_getproduct.Fields.Item("title").Value)%>" data-column="title">
            </div>

            <div class="form-group">
                <label class="font-weight-bold" for="seo_meta_title" data-friendly="SEO title">SEO Unique title (60 Characters Max):</label>
                <input class="form-control form-control-sm" name="seo_meta_title" type="text" id="seo_meta_title" maxlength="60" value="<%=Server.HTMLEncode(rs_getproduct.Fields.Item("seo_meta_title").Value & "") %>" data-column="seo_meta_title">
                <div id="msg-seo-title"></div>
            </div>

            <div class="form-group">
                <label class="font-weight-bold" for="seo_meta_description" data-friendly="SEO title">SEO Description (160 Characters Max):</label>
                <textarea class="form-control form-control-sm" name="seo_meta_description" id="seo_meta_description" rows="3" maxlength="160" data-column="seo_meta_description" data-friendly="Product description" placeholder="Sell product, list gauge range or single gauge"><%= Server.HTMLEncode(rs_getproduct.Fields.Item("seo_meta_description").Value & "") %></textarea>
            </div>


			<div class="form-group">
				<label class="font-weight-bold" for="description">Product Description</label>
				<textarea class="form-control form-control-sm" name="description" id="description" data-column="description" data-friendly="Product description"><%=(rs_getproduct.Fields.Item("description").Value)%></textarea>
			</div>
<!--  TO VIEW CURRENT HS CODES GO TO THIS LINK   https://hts.usitc.gov/current   -->
<%
current_tariff_code = rs_getproduct.Fields.Item("tariff_code").Value

if current_tariff_code = "7117.90.9000" then
	tariff_dropdown = "Any base metal jewelry"
elseif current_tariff_code = "7113.11.5000" then
	tariff_dropdown = "Solid silver jewelry"
elseif current_tariff_code = "7113.19.5090" then
	tariff_dropdown = "Solid gold or platinum jewelry"
elseif current_tariff_code = "7113.20.5000" then
	tariff_dropdown = "Plated gold or platinum jewelry"
elseif current_tariff_code = "3401.30.5000" then
	tariff_dropdown = "Soap cleansers"
elseif current_tariff_code = "3304.99.5000" then
	tariff_dropdown = "Oil / Lotion products"
elseif current_tariff_code = "8203.20.6060" then
	tariff_dropdown = "Tools"
elseif current_tariff_code = "6109.10.0040" then
	tariff_dropdown = "Shirts"
	elseif current_tariff_code = "7116.20.0500" then
	tariff_dropdown = "Stone plugs"
end if
%>
			<div class="form-group">
                <label class="font-weight-bold" for="tariff_code">Harmonized Tariff Code
					<i class="fa fa-lg fa-information text-secondary pointer" data-toggle="modal" data-target="#modal-tariff-info"></i>

				</label>
                <select class="form-control form-control-sm" name="tariff_code" data-column="tariff_code" data-friendly="HS Tariff Code" >
                    <option value="<%= rs_getproduct.Fields.Item("tariff_code").Value %>" selected><%= tariff_dropdown %></option>
					<option value="7117.90.9000">7117.90.9000 - Any base material jewelry</option>
					<option value="7113.11.5000">7113.11.5000 - Base material from solid silver</option>
					<option value="7113.19.5090">7113.19.5090 - Solid gold or platinum jewelry</option>
					<option value="7113.20.5000">7113.20.5000 -  Steel or titanium plated with gold or platinum</option>
					<option value="7116.20.0500">7116.20.0500 - Semiprecious stone plugs (Natural, synthetic, or reconstructed)</option>
					<option value="3401.30.5000">3401.30.5000 - Soap cleansers</option>
					<option value="3304.99.5000">3304.99.5000 - Oil / Lotion products</option>
					<option value="8203.20.6060">8203.20.6060 - Tools</option>
					<option value="6109.10.0040">6109.10.0040 - Shirts</option>
				</select>
			</div>
	<% 
	if (rs_getproduct.Fields.Item("customorder").Value) = "yes" then
		preorder_fields = ""
	else
		preorder_fields = "display:none"
	end if
	%>

	<div class="custom-control custom-checkbox">
		<input name="custom" id="custom" type="checkbox" class="custom-control-input ispreorder" value="yes" <% if (rs_getproduct.Fields.Item("customorder").Value) = "yes" then %>checked<% end if %> data-unchecked="no" data-column="customorder" data-friendly="PRE-ORDER">
		<label class="custom-control-label" for="custom">Is this product a pre-order?</label>
	  </div>
            
		
		
		<div class="preorder-fields" style="<%= preorder_fields %>">

		<h4 class="mt-4">Pre-order options</h4>

			<div class="custom-control custom-checkbox">
				<input class="custom-control-input" name="preorder_nospecs" id="preorder_nospecs" type="checkbox" value="1" <% if (rs_getproduct.Fields.Item("preorder_nospecs").Value) = "1" then %>checked<% end if %> data-unchecked="0" data-column="preorder_nospecs" data-friendly="Preorder no specs">
				<label class="custom-control-label" for="preorder_nospecs">Do not display main specs box</label>
			  </div>

		<div class="form-group mb-1">
		<select name="preorder_field1_label" data-column="preorder_field1_label" data-friendly="Preorder Label #1" class="form-control form-control-sm  preorder_label_selects">
		
			<% if rs_getproduct.Fields.Item("preorder_field1_label").Value <> "" then %>
				<option selected value="<%= rs_getproduct.Fields.Item("preorder_field1_label").Value %>"><%= rs_getproduct.Fields.Item("preorder_field1_label").Value %></option>
			<% else %>
				<option>Select field #1 type:</option>
			<% end if %>					
				<!--#include file="products/inc_preorder_label_select_options.asp" -->
		</select>	
	</div>

		<input class="form-control form-control-sm mb-3" name="preorder_field1" type="text" placeholder="Specs field #1" <% if rs_getproduct.Fields.Item("preorder_field1").Value <> "" then %>value="<%= Server.HTMLEncode(rs_getproduct.Fields.Item("preorder_field1").Value) %>"<% end if %> data-column="preorder_field1" data-friendly="Preorder specs field 1">

		<div class="form-group mb-1">
		<select name="preorder_field2_label" data-column="preorder_field2_label" data-friendly="Preorder Label #2" class="form-control form-control-sm  preorder_label_selects">
		
			<% if rs_getproduct.Fields.Item("preorder_field2_label").Value <> "" then %>
				<option selected value="<%= rs_getproduct.Fields.Item("preorder_field2_label").Value %>"><%= rs_getproduct.Fields.Item("preorder_field2_label").Value %></option>
			<% else %>
				<option>Select field #2 type:</option>
			<% end if %>					
				<!--#include file="products/inc_preorder_label_select_options.asp" -->
		</select>	
	</div>
		
		<input class="form-control form-control-sm mb-3" name="preorder_field2" type="text" placeholder="Specs field #2" <% if rs_getproduct.Fields.Item("preorder_field2").Value <> "" then %>value="<%= Server.HTMLEncode(rs_getproduct.Fields.Item("preorder_field2").Value) %>"<% end if %> data-column="preorder_field2" data-friendly="Preorder specs field 2">
		
		<div class="form-group mb-1">
		<select name="preorder_field3_label" data-column="preorder_field3_label" data-friendly="Preorder Label #3" class="form-control form-control-sm  preorder_label_selects">
		
			<% if rs_getproduct.Fields.Item("preorder_field3_label").Value <> "" then %>
				<option selected value="<%= rs_getproduct.Fields.Item("preorder_field3_label").Value %>"><%= rs_getproduct.Fields.Item("preorder_field3_label").Value %></option>
			<% else %>
				<option>Select field #3 type:</option>
			<% end if %>					
				<!--#include file="products/inc_preorder_label_select_options.asp" -->
		</select>	
	</div>		
		
		<input class="form-control form-control-sm mb-3" name="preorder_field3" type="text" placeholder="Specs field #3" <% if rs_getproduct.Fields.Item("preorder_field3").Value <> "" then %>value="<%= Server.HTMLEncode(rs_getproduct.Fields.Item("preorder_field3").Value) %>"<% end if %> data-column="preorder_field3" data-friendly="Preorder specs field 3">

		<div class="form-group mb-1">
		<select name="preorder_field4_label" data-column="preorder_field4_label" data-friendly="Preorder Label #4" class="form-control form-control-sm  preorder_label_selects">
		
			<% if rs_getproduct.Fields.Item("preorder_field4_label").Value <> "" then %>
				<option selected value="<%= rs_getproduct.Fields.Item("preorder_field4_label").Value %>"><%= rs_getproduct.Fields.Item("preorder_field4_label").Value %></option>
			<% else %>
				<option>Select field #4 type:</option>
			<% end if %>					
				<!--#include file="products/inc_preorder_label_select_options.asp" -->
		</select>	
	</div>
		
		<input class="form-control form-control-sm mb-3" name="preorder_field4" type="text" placeholder="Specs field #4" <% if rs_getproduct.Fields.Item("preorder_field4").Value <> "" then %>value="<%= Server.HTMLEncode(rs_getproduct.Fields.Item("preorder_field4").Value) %>"<% end if %> data-column="preorder_field4" data-friendly="Preorder specs field 4">

		<div class="form-group mb-1">
		<select name="preorder_field5_label" data-column="preorder_field5_label" data-friendly="Preorder Label #5" class="form-control form-control-sm  preorder_label_selects">
		
			<% if rs_getproduct.Fields.Item("preorder_field5_label").Value <> "" then %>
				<option selected value="<%= rs_getproduct.Fields.Item("preorder_field5_label").Value %>"><%= rs_getproduct.Fields.Item("preorder_field5_label").Value %></option>
			<% else %>
				<option>Select field #5 type:</option>
			<% end if %>					
				<!--#include file="products/inc_preorder_label_select_options.asp" -->
		</select>	
	</div>
		
		<input class="form-control form-control-sm mb-3" name="preorder_field5" type="text" placeholder="Specs field #5" <% if rs_getproduct.Fields.Item("preorder_field5").Value <> "" then %>value="<%= Server.HTMLEncode(rs_getproduct.Fields.Item("preorder_field5").Value) %>"<% end if %> data-column="preorder_field5" data-friendly="Preorder specs field 5">

		<div class="form-group mb-1">
		<select name="preorder_field6_label" data-column="preorder_field6_label" data-friendly="Preorder Label #6" class="form-control form-control-sm  preorder_label_selects">
		
			<% if rs_getproduct.Fields.Item("preorder_field6_label").Value <> "" then %>
				<option selected value="<%= rs_getproduct.Fields.Item("preorder_field6_label").Value %>"><%= rs_getproduct.Fields.Item("preorder_field6_label").Value %></option>
			<% else %>
				<option>Select field #6 type:</option>
			<% end if %>					
				<!--#include file="products/inc_preorder_label_select_options.asp" -->
		</select>
	</div>

		<input class="form-control form-control-sm mb-3" name="preorder_field6" type="text" placeholder="Specs field #6	(Not required for customer to fill out)" <% if rs_getproduct.Fields.Item("preorder_field6").Value <> "" then %>value="<%= Server.HTMLEncode(rs_getproduct.Fields.Item("preorder_field6").Value) %>"<% end if %> data-column="preorder_field6" data-friendly="Preorder specs field 6">
		
		<div class="form-group mb-1">
		<select name="preorder_field7_label" data-column="preorder_field7_label" data-friendly="Preorder Label #7" class="form-control form-control-sm  preorder_label_selects">
		
			<% if rs_getproduct.Fields.Item("preorder_field7_label").Value <> "" then %>
				<option selected value="<%= rs_getproduct.Fields.Item("preorder_field7_label").Value %>"><%= rs_getproduct.Fields.Item("preorder_field7_label").Value %></option>
			<% else %>
				<option>Select field #7 type:</option>
			<% end if %>					
				<!--#include file="products/inc_preorder_label_select_options.asp" -->
		</select>
	</div>

		<input class="form-control form-control-sm mb-3" name="preorder_field7" type="text" placeholder="Specs field #7	(Not required for customer to fill out)" <% if rs_getproduct.Fields.Item("preorder_field7").Value <> "" then %>value="<%= Server.HTMLEncode(rs_getproduct.Fields.Item("preorder_field7").Value) %>"<% end if %> data-column="preorder_field7" data-friendly="Preorder specs field 7">
		
		
		
		</div>
</div><!-- end first column -->


    <div class="col-sm px-4 border-right border-left small">  
        <div class="container w-100">
            <div class="row">
                <div class="col-8 p-0 h4">Tags & Filters</div>
            </div>
        </div>    
        

        
        <div class="form-group">

			<div class="container p-0">
				<div class="row">
				  <div class="col-auto">
					<label for="category" class="font-weight-bold">Categories</label> 
				  </div>
				  <div class="col text-right">
					<button class="btn btn-sm btn-secondary py-0 mb-1" id="manage_categories" data-toggle="modal" data-target="#modal-show-categories">Manage categories</button>
				  </div>
				</div>
			</div> 
		
            <select class="select-category" id="select-category" name="category" data-column="jewelry" data-friendly="Categories" multiple>
            
        <% if rs_getproduct.Fields.Item("jewelry").Value <> "" and Instr(rs_getproduct.Fields.Item("jewelry").Value, "null") = 0 then

        ' break full text stored values out into an array to have an <option> selected for each entry
            jewelry_array = split(rs_getproduct.Fields.Item("jewelry").Value," ")
					selected_array = Array()
                    For Each strItem In jewelry_array
                        if strItem <> "" and strItem <> "null " then 

							ReDim Preserve selected_array(UBound(selected_array) + 1)
							selected_array(UBound(selected_array)) = strItem

                        end if 			
                    Next
					selected_array = getUniqueItems(selected_array)
        end if ' if jewelry is not null
                    %>	
					
				<% While NOT rs_getCategories.EOF %>
					<option <% if rs_getproduct("jewelry") <> "" then %><%=IIF(isItemInArray(selected_array, rs_getCategories("category_tag")), "selected ","")%><% end if %> <%=IIF(Instr(rs_getCategories("category_name"), "__") > 0, " disabled", "")%> value="<%= Trim(rs_getCategories("category_tag")) %>"><%= Trim(rs_getCategories("category_name")) %></option>
					<%rs_getCategories.MoveNext()%>
				<% Wend %>	
				

            </select>			

			<div class="custom-control custom-checkbox">
				<input type="checkbox" class="custom-control-input" name="pair" id="pair" value="yes"  <% if (rs_getproduct.Fields.Item("pair").Value) = "yes" then %>checked<% end if %> data-unchecked="no" data-column="pair">
				<label class="custom-control-label" for="pair">Pair</label>
			  </div>

        </div>
		
        <div class="form-group">
            <label class="font-weight-bold" for="piercing_type">Piercing type</label>    
            <select class="select-piercing_type"  name="piercing_type" data-column="piercing_type" data-friendly="Piercing type"  multiple>
            <% if rs_getproduct.Fields.Item("piercing_type").Value <> "" then
            ' break full text stored values out into an array to have an <option> selected for each entry
            type_array = split(rs_getproduct.Fields.Item("piercing_type").Value,"piercing_type:")
                    For Each strItem In type_array
                        if strItem <> "" and strItem <> "null " then 
                        %>
                            <option selected value="<%= strItem %>"><%= strItem %></option>
                        <%
                        end if 			
                    Next
            end if ' if piercing_tye is not null
                    %>		
                    <option value="">___ EAR ____________</option>
                        <option value="Anti-tragus">Anti-tragus</option>
                        <option value="Basic ear piercing">Basic ear piercing</option>
                        <option value="Conch">Conch</option>
                        <option value="Daith">Daith</option>
                        <option value="Helix">Helix</option>
                        <option value="Industrial">Industrial</option>
                        <option value="Lobe">Lobe</option>
                        <option value="Rook">Rook</option>			
                        <option value="Snug">Snug</option>
                        <option value="Stretched lobe">Stretched lobe</option>
                        <option value="Tragus">Tragus</option>
                    
                    <option value="">&nbsp;</option>
                    <option value="">___ FACE ____________</option>
                        <option value="Bites">Bites</option>
                        <option value="Bridge">Bridge</option>
                        <option value="Cheek">Cheek</option>
                        <option value="Eyebrow">Eyebrow</option>
                        <option value="Jestrum">Jestrum</option>
                        <option value="Labret">Labret</option>
                        <option value="Lip">Lip</option>
                        <option value="Philtrum">Philtrum</option>
                        <option value="Vertical labret">Vertical labret</option>
                    
                    <option value="">&nbsp;</option>
                    <option value="">___ GENITAL ____________</option>
                        <option value="Ampallang">Ampallang</option>
                        <option value="Apadravya">Apadravya</option>
                        <option value="Clitoris">Clitoris</option>
                        <option value="Christina">Christina</option>
                        <option value="Dydoe">Dydoe</option>
                        <option value="Foreskin">Foreskin</option>
                        <option value="Fourchette">Fourchette</option>
                        <option value="Frenum">Frenum</option>
                        <option value="Guiche">Guiche</option>
                        <option value="Horizontal hood">Horizontal hood</option>
                        <option value="Labia">Labia</option>
                        <option value="Prince Albert">Prince Albert</option>
                        <option value="Scrotum">Scrotum</option>	
                        <option value="Vertical hood">Vertical hood</option>
                    <option value="">&nbsp;</option>
                    <option value="">___ NOSE ____________</option>
                        <option value="Nostril">Nostril</option>
                        <option value="Septum">Septum</option>
                    
                    <option value="">&nbsp;</option>
                    <option value="">___ OTHER ____________</option>
                    <option value="Microdermal">Microdermal</option>
                    <option value="Navel">Navel</option>
                    <option value="Nipple">Nipple</option>
                    <option value="Surface">Surface</option>
                    <option value="Tongue">Tongue</option>
                    <option value="Tongue web">Tongue web</option>
                    <option value="">&nbsp;&nbsp;&nbsp;&nbsp;</option>
                    <option value="None">None</option>
            </select>
        </div>

        <div class="form-group">
            <label class="font-weight-bold" for="flares">Flare type</label>     
            <select class="select-flares" name="flares" data-column="flare_type" data-friendly="Plug flares"  multiple>
                <% if rs_getproduct.Fields.Item("flare_type").Value <> "" then

                ' break full text stored values out into an array to have an <option> selected for each entry
                flare_array = split(rs_getproduct.Fields.Item("flare_type").Value," , ")
                        For Each strItem In flare_array
                            if strItem <> "" and strItem <> "null " then 
                            %>
                                <option selected value="<%= strItem %>"><%= strItem %></option>
                            <%
                            end if 			
                        Next
                end if ' if threading is not null
                %>	 
                <option value="Single flare">Single flare</option>
                <option value="Double flare">Double flare</option>
                <option value="No flare">No flare</option>
                <option value="Screw on">Screw/thread on back</option>
            </select>
		</div>
		
        <div class="form-group">
            <label class="font-weight-bold" for="threading">Threading</label>  
            <select class="select-threading" name="threading" data-column="internal" data-friendly="Threading type"  multiple>
            <% if rs_getproduct.Fields.Item("internal").Value <> "" and Instr(rs_getproduct.Fields.Item("internal").Value, "null") = 0 then

            ' break full text stored values out into an array to have an <option> selected for each entry
                threading_array = split(rs_getproduct.Fields.Item("internal").Value," , ")
                        For Each strItem In threading_array
                            if strItem <> "" and strItem <> "null " then 
                            %>
                                <option selected value="<%= strItem %>"><%= strItem %></option>
                            <%
                            end if 			
                        Next
            end if ' if threading is not null
                    %>	
                    
                <option value="Externally threaded">Externally threaded</option>
                <option value="Internally threaded">Internally threaded</option>
				<option value="Threadless">Threadless</option>
				<option value="Push pin">Push pin</option>
            </select>
        </div>

        <div class="form-group">
			<div class="container p-0">
				<div class="row">
				  <div class="col-auto">
					<label class="font-weight-bold align-bottom" for="materials_main">Materials</label>
				  </div>
				  <div class="col text-right">
					<button class="btn btn-sm btn-secondary py-0 mb-1 mr-1 btn-clear-fields" id="clear-variant-materials">Clear all</button>
					<button class="btn btn-sm btn-secondary py-0 mb-1" id="apply_all_material">Copy materials to variants</button>
					<button class="btn btn-sm btn-secondary py-0 mb-1" id="manage_materials" data-toggle="modal" data-target="#modal-show-materials">Manage materials</button>
				  </div>
				</div>
			</div>
            
            <select class="" name="materials_main" id="materials_main" data-column="material" data-friendly="Materials"  multiple>
                <% if rs_getproduct.Fields.Item("material").Value <> "" then
					' break full text stored values out into an array to have an <option> selected for each entry

					selected_materials_array = getUniqueItems(split(rs_getproduct.Fields.Item("material").Value,","))
                end if 

                %>	
				<% While NOT rs_getMaterials.EOF %>
					<option <% if rs_getproduct("material") <> "" then %><%=IIF(isItemInArray(selected_materials_array, rs_getMaterials("material_name")), "selected ","")%><% end if %> <%=IIF(Instr(rs_getMaterials("material_name"), "__") > 0, " disabled", "")%> value="<%= Trim(rs_getMaterials("material_name")) %>"><%= Trim(rs_getMaterials("material_name")) %></option>
					<%rs_getMaterials.MoveNext()%>
				<% Wend %>			
            
            </select>
        </div>

        <div class="form-group">
			<div class="container p-0">
				<div class="row">
				  <div class="col-auto">
					<label class="font-weight-bold" for="wearable_main">Wearable</label>
				  </div>
				  <div class="col text-right">
					<button class="btn btn-sm btn-secondary py-0 mb-1 mr-1 btn-clear-fields" id="clear-variant-wearable">Clear all</button>
					<button class="btn btn-sm btn-secondary py-0 mb-1" id="apply_all_wearable_materials">Copy wearable to variants</button>
					<button class="btn btn-sm btn-secondary py-0 mb-1" id="manage_wearable" data-toggle="modal" data-target="#modal-show-materials">Manage materials</button>
				  </div>
				</div>
			  </div>
            
				<select class="form-control form-control-sm " name="wearable_main" id="wearable_main">
					<option>Select wearable material...</option>					
					<% While NOT rs_getWearableMaterials.EOF %>
					<option <%=IIF(Instr(rs_getWearableMaterials("material_name"), "__") > 0, " disabled", "")%> value="<%= Trim(rs_getWearableMaterials("material_name")) %>"><%= Trim(rs_getWearableMaterials("material_name")) %></option>
					<% 
						rs_getWearableMaterials.MoveNext()
						Wend
					%> 	
				</select>	
        </div>

        <div class="form-group">
			<div class="container p-0">
				<div class="row">
				  <div class="col-auto">
					<label class="font-weight-bold" for="colors_main">Colors</label>
				  </div>
				  <div class="col text-right">
					<button class="btn btn-sm btn-secondary py-0 mb-1 mr-1 btn-clear-fields" id="clear-variant-colors">Clear all</button>
					<button class="btn btn-sm btn-secondary py-0 mb-1" id="apply_all_colors">Copy colors to variants</button>
				  </div>
				</div>
			  </div>
           
				<select class="select-colors"  name="colors_main" id="colors_main" data-column="colors_main"   multiple>
					<option>Select colors...</option>					
				<% for each x in color_array %>
					<option value="<%= x %>"><%= x %></option>
				<% next %>
				</select>	
		</div>
		
		<div class="form-group">
			  
			<div class="container p-0">
				<div class="row">
				  <div class="col-auto">
					<label class="font-weight-bold" for="tags">Tags:</label>  
				  </div>
				  <div class="col text-right">
					<button class="btn btn-sm btn-secondary py-0 mb-1" id="manage_tags" data-toggle="modal" data-target="#modal-show-tags">Manage tags</button>
				  </div>
				</div>
			</div>			
			<select class="select-tags" id="select-tags" name="tags" data-column="tags" data-friendly="Tags" multiple>
		
			<% if rs_getproduct.Fields.Item("tags").Value <> "" and Instr(rs_getproduct.Fields.Item("tags").Value, "null") = 0 then
		
			 ' break full text stored values out into an array to have an <option> selected for each entry
				tags_selected_array = split(rs_getproduct.Fields.Item("tags").Value," ")
						For Each strItem In tags_selected_array
							if strItem <> "" and strItem <> "null " then 
							%>
								<option selected value="<%= strItem %>"><%= strItem %></option>
							<%
							end if 			
						Next
			end if
			%>	
		
            <% While NOT rs_getTags.EOF %>
                <option value="<%=(rs_getTags.Fields.Item("tag").Value)%>"><%=(rs_getTags.Fields.Item("tag").Value)%>                </option>
            <% 
                rs_getTags.MoveNext()
                Wend
            %> 
			</select>
			</div>

			
		<div class="form-group">
			<label class="font-weight-bold" for="vartype">Status</label>
			<select class="form-control form-control-sm " name="vartype" data-column="type" data-friendly="Status">
				<option>None</option>
				<option value="Clearance">Clearance</option>
				<option value="limited">Limited</option>
				<option value="Discontinued">Discontinued</option>
				<option value="One time buy">One time buy</option>
				<option value="Consignment">Consignment</option>
				<option value="<%=(rs_getproduct.Fields.Item("type").Value)%>" selected><%=(rs_getproduct.Fields.Item("type").Value)%></option>
			</select>
		</div>


        <div class="form-group">
            <label class="font-weight-bold" for="brand_name">Brand</label>
            <select class="form-control form-control-sm " name="brand_name" data-column="brandname" data-friendly="Brands" >
                <option value="<%=(rs_getproduct.Fields.Item("brandname").Value)%>" selected><%=(rs_getproduct.Fields.Item("brandname").Value)%></option>
                <% 
                While NOT rs_getbrand.EOF 
                %>
                <option value="<%=(rs_getbrand.Fields.Item("name").Value)%>"><%=(rs_getbrand.Fields.Item("name").Value)%>
                </option>
                <% 
                rs_getbrand.MoveNext()
                Wend
                %>                
                <option value="None">None</option>
            </select>
        </div>

		<div class="form-group">
            <label class="font-weight-bold" for="country_origin">Country of origin</label>
            <select class="form-control form-control-sm " name="country_origin" data-column="country_origin" data-friendly="Origin Country" >
                <option value="<%= rs_getproduct("country_origin") %>" selected><%= rs_getproduct.Fields.Item("country_origin") %></option>
                <% 
                While NOT rs_getOriginCountries.EOF 
                %>
                <option value="<%= rs_getOriginCountries("country") %>"><%= rs_getOriginCountries("country") %>
                </option>
                <% 
                rs_getOriginCountries.MoveNext()
                Wend
                %>                
                <option value="">None</option>
            </select>
        </div>						  
		<div class="custom-control custom-checkbox">
			<input class="custom-control-input" name="autoclavable" id="autoclavable" type="checkbox" value="1" <% if rs_getproduct.Fields.Item("autoclavable").Value = 1 then %>checked<% end if %> data-unchecked="0" data-column="autoclavable" data-friendly="Autoclavable">
			<label class="custom-control-label" for="autoclavable">Autoclavable?</label>
		  </div>  

		  <div class="custom-control custom-checkbox">
			<input class="custom-control-input" name="pinned_product" id="pinned_product" type="checkbox" value="1" <% if (rs_getproduct.Fields.Item("pinned_product").Value) = "1" then %>checked<% end if %> data-unchecked="0" data-column="pinned_product" data-friendly="Pin product">
			<label class="custom-control-label" for="pinned_product">Pin product to top of search results</label>
		  </div>      

    </div><!-- end second column -->
    

    <div class="col-sm pl-4 small"> 
            <div class="container w-100">
                    <div class="row">
						<div class="col-4 p-0 h4">Photos</div>
						<div class="col-8 p-0 text-right">
							<button class="btn btn-sm btn-secondary mt-2" type="button" id="combine_reviews">Combine customer reviews & photos</button>
					</div>
                    </div>
				</div>  

				<div class="card bg-light p-3 my-2" style="display:none" id="combine_div" role="alert">
					This will move the customer submitted reviews & photos FROM the product ID entered below into this current listing
					<div class="form-inline mt-2">
						<input class="form-control form-control-sm" type="text" name="id-transfer-reviews" id="id-transfer-reviews" placeholder="Product ID to transfer">
					
						<button class="btn btn-sm btn-secondary ml-3" type="button" id="combine-submit">Submit</button>
					</div>
					<div class="alert alert-success mt-3" style="display:none" id="combine_success">
						Transfer successful
					</div>
				</div><!-- combine -->
				
				<div class="card bg-light p-3 mt-3" role="alert">
					<div id="img_remove"></div>
					<div class="form-group">
						<div class="form-check form-check-inline">
						  <input class="form-check-input" type="radio" name="phototype" id="opt-main-image" value="opt-main-image" checked>
						  <label class="form-check-label" for="opt-main-image">Main Image</label>
						</div>
						<div class="form-check form-check-inline">
						  <input class="form-check-input" type="radio" name="phototype" id="opt-additional-image" value="opt-additional-image">
						  <label class="form-check-label" for="opt-additional-image">Additional Image</label>
						</div>	
						<div class="form-check form-check-inline">
						  <input class="form-check-input" type="radio" name="phototype" id="opt-video" value="video">
						  <label class="form-check-label" for="opt-video">Video</label>
						</div>							
					</div>
										
					<div class="form-group">
						<div id="dropzone-area">
						<form action="#" class="dropzone needsclick dz-clickable" id="frmUpload">
						  <div class="dz-message needsclick">
							<button type="button" class="dz-button">Drop <span>MAIN</span> images here or click to upload.</button><br>
							<span class="note needsclick">Upload all 3 image sizes together<br>
								1000 x 1000, 400 x 400, and 90 x 90</span>
						  </div>

						</form>
						</div>						
					</div>
					<div id="img_description" class="form-group d-none">
					   <input class="form-control form-control-sm" type="text" name="add_img_description" id="add_img_description" placeholder="Color / Description" maxlength="50">
					</div>
					<div class="d-inline-flex w-100 justify-content-between">
						<button class="btn btn-sm btn-secondary" type="button" id="clear_dropzone"> Clear </button>	
						<button class="btn btn-sm btn-secondary w-100 ml-2" type="button" id="btn-upload" data-productid="<%= rs_getproduct.Fields.Item("ProductID").Value %>"> Upload Images</button>
					</div>   
				</div>
				
				
				<div class="form-group mt-2">
					<label class="font-weight-bold" for="color-charts">Color chart(s)&nbsp;&nbsp;<a href="ColorCharts.html" target="_new">HTML links</a></label>
					<textarea class="form-control form-control-sm" name="color-charts" data-column="ColorChart" data-friendly="Color charts"><%=(rs_getproduct.Fields.Item("ColorChart").Value)%></textarea>
				</div>

             
                
           
				<div class="container w-100">
                    <div class="row">
                        <div class="col-8 p-0 h4">Sales</div>
                    </div>
				</div>  

               <div class="form-group">
                <label class="font-weight-bold" for="discount">Discount</label>

                <select class="form-control form-control-sm" name="discount" data-column="SaleDiscount" data-friendly="Discount amount" >
                    <option value="<%= (rs_getproduct.Fields.Item("SaleDiscount").Value) %>" selected><% if rs_getproduct.Fields.Item("SaleDiscount").Value = 0 then %>None<%else%><%= (rs_getproduct.Fields.Item("SaleDiscount").Value) %><% end if %></option>
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

			<%

'====== CONFIGURE DATE TO SHOW CORRECTLY IN Field
if rs_getproduct.Fields.Item("sale_expiration").Value <> "" then
var_sale_expiration = rs_getproduct.Fields.Item("sale_expiration").Value
	var_sale_expiration = DatePart("yyyy",var_sale_expiration) _
	& "-" & Right("0" & DatePart("m",var_sale_expiration), 2) _
	& "-" & Right("0" & DatePart("d",var_sale_expiration), 2)
end if
'======== SALE DISCOUNTS ARE SET BACK TO 0 VIA A DAILY JOB IN THE DATABASE THAT CHECKS THE CURRENT EXPIRATION ON ALL PRODUCTS
						  
	  
				   
		
%>
  
			<div class="form-group">
                <label class="font-weight-bold" for="sale-expiration">Sale expiration date</label>
                <input class="form-control form-control-sm" name="sale-expiration" type="date" id="sale-expiration" value="<%= var_sale_expiration %>" data-column="sale_expiration">
																								 
            </div>

                <% 	if (rs_getproduct.Fields.Item("SaleExempt").Value) = 1 then
                    var_checked = "checked"
                else
                    var_checked = ""
                end if
				%>
				<div class="custom-control custom-checkbox">
					<input class="custom-control-input" name="exempt" id="exempt" type="checkbox" value="1" data-column="SaleExempt" <%= var_checked %> data-unchecked="0" data-column="SaleExempt" data-friendly="Sale exempt">
					<label class="custom-control-label" for="exempt">Sale exempt</label>
				  </div>  

                <% 	if rs_getproduct.Fields.Item("secret_sale").Value = 1 then
                    var_checked = "checked"
                else
                    var_checked = ""
                end if
				%>
				<div class="custom-control custom-checkbox">
					<input class="custom-control-input" name="secret_sale" id="secret_sale" type="checkbox" value="1" data-column="secret_sale" <%= var_checked %> data-unchecked="0" data-column="secret_sale" data-friendly="Secret sale">
					<label class="custom-control-label" for="secret_sale">Secret Sale</label>
				  </div> 

               
				  <div class="form-group mt-3">
					<label class="h4" for="notes">Private notes</label>
					<textarea class="form-control form-control-sm" rows="10" name="notes" data-column="ProductNotes" data-friendly="Private notes"><%=(rs_getproduct.Fields.Item("ProductNotes").Value)%></textarea>
				</div>                
		</div><!-- 3rd column-->
    </div><!-- row -->
    </div><!-- container -->

	<div class="mt-5 ml-1 mb-2 form-inline ajax-update">
		<a class="btn btn-sm btn-secondary mr-5" href="print-friendly-product.asp?ProductID=<%= request.querystring("ProductID") %>" target="_blank">Print friendly</a>
		<% 	if rs_getproduct.Fields.Item("to_be_pulled").Value = 1 then
		var_checked = "checked"
		else
		var_checked = ""
		end if
		%>
		<span class="custom-control custom-checkbox">
			<input class="custom-control-input" name="to_be_pulled" id="to_be_pulled" type="checkbox" value="1" data-column="to_be_pulled" <%= var_checked %> data-unchecked="0" data-column="to_be_pulled" data-friendly="To be pulled">
			<label class="custom-control-label" for="to_be_pulled">To be pulled</label>
		</span> 
</div>
<div class="loader-div" style="display:none"></div>


<table class="table table-sm table-borderless small css-product-edit " id="details-table">
<thead class="thead-dark text-center">
	<tr>
		<th class="sticky-top" scope="col">&nbsp;</th>
		<th class="sticky-top" scope="col">Sort</th>
		<th class="sticky-top" scope="col">Section</th>
		<th class="sticky-top" scope="col">Location</th>
		<th class="sticky-top" scope="col">Qty</th>
		<th class="sticky-top" scope="col">Max</th>
		<th class="sticky-top" scope="col">Thresh</th>
		<th class="sticky-top" scope="col">Gauge</th>
		<th class="sticky-top" scope="col">Length</th>
		<th class="sticky-top" scope="col">Details</th>
		<th class="sticky-top" scope="col">Retail</th>
		<th class="sticky-top" scope="col">Wlsl</th>
		<th class="sticky-top" scope="col">
		<button class="btn btn-sm btn-secondary sticky-top applyall" type="button" data-column="sku" data-field="sku">Apply to all</button>
		Vendor SKU
		</th>
		<th class="sticky-top" scope="col">Active</th>
		<th class="sticky-top" scope="col">Last sold</th>
		<th class="sticky-top" scope="col">Copy/Move</th>
	</tr>
</thead>

<form id="add-detail">
	<tr class="add-new">
		<td>&nbsp;</td>
		<td>
			<input class="form-control form-control-sm"  style="width: 50px" name="sort" type="text" value="0">
		</td>
		<td>
			<select class="form-control form-control-sm w-auto" name="section" id="add-section" class="">
				<% While NOT rs_getsections.EOF %>                          
				<option value="<%=(rs_getsections.Fields.Item("ID_Number").Value)%>"><%=(rs_getsections.Fields.Item("ID_Description").Value)%></option>

				<% 
				rs_getsections.MoveNext()
				Wend
				rs_getsections.MoveFirst()
				%> 
			</select>
		</td>
		<td>
			<input class="form-control form-control-sm" style="width: 50px" name="location" type="text" value="0">
		</td>
		<td>
			<input class="form-control form-control-sm check-wholesale" style="width: 50px;border:2px solid #1c9923!important" name="qty-onhand" type="text" value="0" data-pricecheck="new">
		</td>
		<td>
			<input class="form-control form-control-sm" style="width: 50px" name="max" type="number" value="0" min="1" required>
		</td>
		<td>
			<input class="form-control form-control-sm" style="width: 50px" name="thresh" type="number" value="0">
		</td>
		<td>
			<select class="form-control form-control-sm w-auto" name="gauge" class="">
				<option value="" selected>Gauge</option>
				<% While NOT rsGetGauges.EOF %>
				<option value="<%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %>"><%= rsGetGauges.Fields.Item("GaugeShow").Value %></option>
				<% 
				rsGetGauges.MoveNext()
				Wend 
				rsGetGauges.ReQuery() %>
				<option value="">None</option>
			</select>
		</td>
		<td>
			<select class="form-control form-control-sm w-auto" name="length" class="">
				<option value="" selected>Length</option>
				<option value="4mm">4mm</option>
				<option value="5mm">5mm</option>
				<option value="3/16&quot;">3/16&quot;</option>
				<option value="1/4&quot;">1/4&quot;</option>
				<option value="7mm">7mm</option>
				<option value="9/32&quot;">9/32&quot;</option>
				<option value="5/16&quot;">5/16&quot;</option>
				<option value="11/32&quot; (9mm)">11/32&quot; (9mm)</option>
				<option value="3/8&quot; (9.5mm)">3/8&quot; (9.5mm)</option>
				<option value="10mm">10mm</option>
				<option value="13/32&quot;">13/32&quot;</option>
				<option value="7/16&quot;">7/16&quot;</option>
				<option value="11mm">11mm</option>
				<option value="12mm">12mm</option>
				<option value="1/2&quot;">1/2&quot;</option>
				<option value="13mm">13mm</option>
				<option value="9/16&quot;">9/16&quot;</option>
				<option value="15mm">15mm</option>
				<option value="5/8&quot;">5/8&quot;</option>
				<option value="11/16&quot;">11/16&quot;</option>
				<option value="18mm">18mm</option>
				<option value="3/4&quot;">3/4&quot;</option>
				<option value="13/16&quot;">13/16&quot;</option>
				<option value="7/8&quot;">7/8&quot;</option>
				<option value="15/16&quot;">15/16&quot;</option>
				<option value="1&quot;">1&quot;</option>
				<option value="1-1/16&quot;">1-1/16&quot;</option>
				<option value="1-1/8&quot;">1-1/8&quot;</option>
				<option value="1-3/16&quot;">1-3/16&quot;</option>
				<option value="1-1/4&quot;">1-1/4&quot;</option>
				<option value="1-5/16&quot;">1-5/16&quot;</option>
				<option value="1-3/8&quot;">1-3/8&quot;</option>
				<option value="1-7/16&quot;">1-7/16&quot;</option>
				<option value="1-1/2&quot;">1-1/2&quot;</option>
				<option value="1-9/16&quot;">1-9/16&quot;</option>
				<option value="1-5/8&quot;">1-5/8&quot;</option>
				<option value="1-11/16&quot;">1-11/16&quot;</option>
				<option value="1-3/4&quot;">1-3/4&quot;</option>
				<option value="1-7/8&quot;">1-7/8&quot;</option>
				<option value="2&quot;">2&quot;</option>
				<option value="2-1/4&quot;">2-1/4&quot;</option>
				<option value="2-1/2&quot;">2-1/2&quot;</option>
				<option value="3&quot;">3&quot;</option>
			</select>
		</td>
		<td>
			<input class="form-control form-control-sm"  style="width: 300px" name="detail" type="text" placeholder="More item details">
		</td>
		<td>
			<input class="form-control form-control-sm check-wholesale pricecheck_retail_new" style="width: 75px" name="retail" type="number" value="0" min="1" step="any" required data-pricecheck="new">
		</td>
		<td>
			<input class="form-control form-control-sm check-wholesale pricecheck_wlsl_new" style="width: 75px" name="wholesale" type="number" value="0" min="1" step="any" data-pricecheck="new">
		</td>
		<td>
			<input class="form-control form-control-sm" style="width: 125px" name="sku" id="sku" type="text" placeholder="Vendor SKU">
		</td>
		<td class="text-center">
			<div class="custom-control custom-checkbox">
				<input class="custom-control-input" name="active" id="active" type="checkbox" value="1" checked>
				<label class="custom-control-label" for="active">
					&nbsp;
				</label>
			  </div> 
		</td>
		<td>
			<button type="submit" class="btn btn-sm btn-primary" id="add-button">Add</button>
		</td>
		<td>
			<span style="display:none" id="move-copy-productid"><span id="move-copy-text"></span> to product # <input  class="form-control form-control-sm" type="text" size="10" name="toggle-productid" id="toggle-productid"></span>
			<input name="date-added" type="hidden" value="<%= date() %>">
			<input name="productid" type="hidden" value="<%= Request.QueryString("ProductID") %>">
		</td>
	</tr>
	<tr>
		<td></td>
		<td class="pb-5" colspan="16">
			<div class="form-inline">
				<span class="mr-1">Weight</span>                     
				<input class="form-control form-control-sm mr-4" style="width: 50px" name="weight" type="text" value="0">
						
				<select class="form-control form-control-sm mr-4 select-detail-wearable-materials" name="wearable_add" id="wearable_add">
					<option>Select wearable material...</option>
					<%rs_getWearableMaterials.MoveFirst()%>
					<% While NOT rs_getWearableMaterials.EOF %>
					<option <%=IIF(Instr(rs_getWearableMaterials("material_name"), "__") > 0, " disabled", "")%> value="<%= Trim(rs_getWearableMaterials("material_name")) %>"><%= Trim(rs_getWearableMaterials("material_name")) %></option>
					<% 
						rs_getWearableMaterials.MoveNext()
						Wend
					%> 
				</select>	
		
			
			<span class="mr-4" style="width:300px">
				<select name="materials_add" id="materials_add" class="select-detail-materials " multiple>
				<%rs_getMaterials.MoveFirst()%>
				<% While NOT rs_getMaterials.EOF %>
					<option <%=IIF(Instr(rs_getMaterials("material_name"), "__") > 0, " disabled", "")%> value="<%= Trim(rs_getMaterials("material_name")) %>"><%= Trim(rs_getMaterials("material_name")) %></option>
					<%rs_getMaterials.MoveNext()%>
				<% Wend %>					
				</select>
			</span>
			
			<span class="mr-4" style="width:250px">
				<select class="" name="colors_add" id="colors_add" class="select-colors " multiple>
					<option>Select colors...</option>
					<% for each x in color_array %>
						<option value="<%= x %>"><%= x %></option>
					<% next %>
				</select>
			</span>
		
		</div>
		</td>
	</tr>
</form>

<form id="frm_filters" action="" method="get">
<input type="hidden" name="ProductID" value="<%=(rs_getproduct.Fields.Item("ProductID").Value)%>" />
	<tr class="bg-secondary" id="filters">
		<td class="py-2 text-white" colspan="5">
			<button class="btn btn-sm btn-dark mr-3 expand-all" type="button"><i class="fa fa-lg fa-angle-double-down" id="btn-expand-all"></i></button>
			<h5 class="d-inline pr-5">
				<%= var_total_details %> total records
			</h5>
			<!--#include file="products/inc_paging.asp" -->
		</td>
	<td class="py-2 text-right" colspan="12">
		<div class="form-inline"style="display: flex; justify-content: flex-end">
			<select class="form-control form-control-sm mr-3" name="filter_gauge" id="filter_gauge">
				<% if request("filter_gauge") <> "" then %>
					<option value="<%= Server.HTMLEncode(request("filter_gauge")) %>">Filtered gauge: <%= request("filter_gauge") %></option>
				<% end if %>
				<option value="">Filter by gauge (Show all)</option>	
				<option value="20g">20g &amp; 18g</option>	
				<option value="16g">16g</option>				
				<option value="14g">14g &amp; 14g/12g</option>	
				<option value="12g">12g</option>
				<option value="10g">10g</option>				
				<option value="8g">8g</option>
				<option value="6g">6g</option>
				<option value="4g">4g</option>
				<option value="2g">2g</option>
				<option value="0g">0g</option>
				<option value="00g">00g</option>
				<option value="7/16&quot;">7/16"</option>
				<option value="1/2&quot;">1/2"</option>
				<option value="9/16&quot;">9/16"</option>
				<option value="5/8&quot;">5/8"</option>				
				<option value="3/4&quot;">3/4"</option>
				<option value="7/8&quot;">7/8"</option>
				<option value="1&quot;">1"</option>
				<option value="odd_small">Odd sizes: 13g/11g/7g/5g/3g/1g</option>
				<option value="odd_large">Odd sizes: 11/16", 13/16", 15/16"</option>
				<option value="odd mm above 00g">odd mm above 00g</option>
				<option value="Between 1 inch - 2 inch">Between 1 inch - 2 inch</option>
				<option value="2 inch - 3 inch">2 inch - 3 inch</option>		
			</select>
			
			<select class="form-control form-control-sm mr-3" name="filter_active" id="filter_active">
				<option value=""><%= filter_select_text %></option>
				<option value="">Show active & inactive</option>
				<option value="active">Show active only</option>
				<option value="inactive">Show inactive only</option>
			</select>
			
			<input class="form-control form-control-sm" type="text" name="filter_detailid" placeholder="Detail ID #" />
		</div>
<%
	'Set limits on how many records are returned without filtering
	if var_total_details > 50 and request("filter_active") = "" then
		var_limit_details = 50
	end if
	if var_total_details <= 50 and var_total_active_details <= 50 and var_total_inactive_details <= 50 then
		var_limit_details = 50
	end if	
	if var_total_active_details > 50 and request("filter_active") = "active" then
		var_limit_details = 50
	end if	
	if var_total_inactive_details > 50 and request("filter_active") = "inactive" then
		var_limit_details = 50
	end if	
	if request("filter_active") <> "" and request("filter_gauge") <> "" then
		var_limit_details = 500
	end if		
 %>

		</td>
	</tr>		
</form>

<tbody id="display-new-row">
</tbody>
<% 	
	var_inactive = 0
	var_inactive_noloop = "no"
	detail_loop_count = 0
	
	if NOT rs_getdetails.EOF then
		rs_getdetails.AbsolutePage = intPage '======== PAGING
	'======== PAGING
	For intRecord = 1 To rs_getdetails.PageSize 
	
	
'	While NOT rs_getdetails.EOF AND detail_loop_count <= var_limit_details 
		detail_loop_count = detail_loop_count + 1
		
	
		
	if rs_getdetails.Fields.Item("active").Value = 0 then
		inactive_class = "table-secondary"
		inactive_field = "background-color:#d6d8db;border-color:#A4A4A4;"
		var_inactive = 1
	else	
		inactive_class = "tbody-active-details"
		inactive_field = ""
		clone = ""
	end if
	%>
	<% if var_inactive = 1 and var_inactive_noloop = "no" then %>
	<tr><td class="text-center py-3" colspan="17"><h4>INACTIVE ITEMS</h4></td></tr>
	<% 
	var_inactive_noloop = "yes"
	end if %>
<tbody class="row-group ajax-update <%= inactive_class %>" id="tbody-<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
	<tr class="show-less details-border detail-main-<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		<td class="align-middle text-nowrap">
			
			<span id="img_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
			<% if rs_getdetails.Fields.Item("img_id").Value <> 0 then %>
				<img style="width:40px;height:auto" src="http://bodyartforms-products.bodyartforms.com/<%= rs_getdetails.Fields.Item("img_thumb").Value %>" class="mini-thumb img_<%= rs_getdetails.Fields.Item("img_id").Value %>" data-name="<%= rs_getdetails.Fields.Item("img_full").Value %>" />
			<% end if %>
			</span>
			<button class="btn btn-sm btn-secondary py-0 px-1 assign_img" type="button"  data-id="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
			<i class="fa fa-image fa-lg"></i></button>
			<button class="btn btn-sm btn-secondary py-0 px-1 view-edits-log" type="button"  data-toggle="modal" data-target="#modal-edits-log"  data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
				<i class="fa fa-table-edit fa-lg"></i></button>
		
			<button class="btn btn-sm btn-secondary px-1 py-0 ml-1 mr-3 expand-one" type="button" data-id="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"><i class="fa fa-angle-double-down" id="expand_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"></i></button>
			
			<span>
				<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>
				<% if rs_getdetails.Fields.Item("BinNumber_Detail").Value <> 0 then %>
				BIN <%=(rs_getdetails.Fields.Item("BinNumber_Detail").Value)%>
				<% end if %>
			</span>	
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm" style="<%= inactive_field %>" name="sort_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= (rs_getdetails.Fields.Item("item_order").Value)%>"  data-column="item_order" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Sort order">
		</td>
		<td class="align-middle">
			<select class="form-control form-control-sm" style="<%= inactive_field %>" name="section_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"  data-column="DetailCode" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Section" class="">
				<option value="<%=(rs_getdetails.Fields.Item("DetailCode").Value)%>" selected>
					<%=(rs_getdetails.Fields.Item("ID_Description").Value)%>
				</option>
				<% While NOT rs_getsections.EOF %>
				<option value="<%=(rs_getsections.Fields.Item("ID_Number").Value)%>">
					<%=(rs_getsections.Fields.Item("ID_Description").Value)%>
				</option>
				<% 
				rs_getsections.MoveNext()
				Wend
				rs_getsections.MoveFirst()
				%> 
			</select>
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm" style="<%= inactive_field %>" name="location_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= (rs_getdetails.Fields.Item("location").Value)%>" data-column="location" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Location">
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm origqty check-wholesale" style="<%= inactive_field %> border:2px solid #1c9923!important" name="qty-onhand_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%=(rs_getdetails.Fields.Item("qty").Value)%>" data-column="qty" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-origqty="<%=(rs_getdetails.Fields.Item("qty").Value)%>" data-friendly="Qty" data-pricecheck="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm" style="<%= inactive_field %>" name="max_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text"value="<% if Not isNull(rs_getdetails.Fields.Item("stock_qty").Value) then%><%=(rs_getdetails.Fields.Item("stock_qty").Value)%><% else %>0<% end if %>" data-column="stock_qty" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Max qty">
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm" style="<%= inactive_field %>" name="thresh_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%=(rs_getdetails.Fields.Item("restock_threshold").Value)%>" data-column="restock_threshold" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Threshold">
		</td>
		<td class="align-middle">
			<select class="form-control form-control-sm" style="<%= inactive_field %>" name="gauge_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="Gauge" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Gauge" class="">
				<option value="<% If (rs_getdetails.Fields.Item("Gauge").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Gauge").Value)%><% end if %>" selected><% If (rs_getdetails.Fields.Item("Gauge").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Gauge").Value)%><% end if %></option>  
				<% While NOT rsGetGauges.EOF %>
				<option value="<%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %>"><%= rsGetGauges.Fields.Item("GaugeShow").Value %></option>
				<% rsGetGauges.MoveNext()
				Wend 
				rsGetGauges.ReQuery() %>
				<option value="">None</option>
			</select>
		</td>
		<td class="align-middle">
			<select class="form-control form-control-sm" style="<%= inactive_field %>" name="length_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="Length" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Length" class="">
				<option value="<% If (rs_getdetails.Fields.Item("Length").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Length").Value)%><% end if %>" selected><% If (rs_getdetails.Fields.Item("Length").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Length").Value)%><% end if %></option>  
				<option value="">None</option>
				<option value="4mm">4mm</option>
				<option value="5mm">5mm</option>
				<option value="3/16&quot;">3/16&quot;</option>
				<option value="1/4&quot;">1/4&quot;</option>
				<option value="7mm">7mm</option>
				<option value="9/32&quot;">9/32&quot;</option>
				<option value="5/16&quot;">5/16&quot;</option>
				<option value="11/32&quot; (9mm)">11/32&quot; (9mm)</option>
				<option value="3/8&quot; (9.5mm)">3/8&quot; (9.5mm)</option>
				<option value="10mm">10mm</option>
				<option value="13/32&quot;">13/32&quot;</option>
				<option value="7/16&quot;">7/16&quot;</option>
				<option value="11mm">11mm</option>
				<option value="12mm">12mm</option>
				<option value="1/2&quot;">1/2&quot;</option>
				<option value="9/16&quot;">9/16&quot;</option>
				<option value="5/8&quot;">5/8&quot;</option>
				<option value="11/16&quot;">11/16&quot;</option>
				<option value="3/4&quot;">3/4&quot;</option>
				<option value="13/16&quot;">13/16&quot;</option>
				<option value="7/8&quot;">7/8&quot;</option>
				<option value="15/16&quot;">15/16&quot;</option>
				<option value="1&quot;">1&quot;</option>
				<option value="1-1/16&quot;">1-1/16&quot;</option>
				<option value="1-1/8&quot;">1-1/8&quot;</option>
				<option value="1-3/16&quot;">1-3/16&quot;</option>
				<option value="1-1/4&quot;">1-1/4&quot;</option>
				<option value="1-5/16&quot;">1-5/16&quot;</option>
				<option value="1-3/8&quot;">1-3/8&quot;</option>
				<option value="1-7/16&quot;">1-7/16&quot;</option>
				<option value="1-1/2&quot;">1-1/2&quot;</option>
				<option value="1-9/16&quot;">1-9/16&quot;</option>
				<option value="1-5/8&quot;">1-5/8&quot;</option>
				<option value="1-11/16&quot;">1-11/16&quot;</option>
				<option value="1-3/4&quot;">1-3/4&quot;</option>
				<option value="1-7/8&quot;">1-7/8&quot;</option>
				<option value="2&quot;">2&quot;</option>
				<option value="2-1/4&quot;">2-1/4&quot;</option>
				<option value="2-1/2&quot;">2-1/2&quot;</option>
				<option value="3&quot;">3&quot;</option>
			</select>
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm" style="<%= inactive_field %>" name="details_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" <% if rs_getdetails.fields.item("ProductDetail1").value <> "" then%>value="<%= Server.HTMLEncode(rs_getdetails.Fields.Item("ProductDetail1").Value)%>"<% end if %> data-column="ProductDetail1" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Extra details">
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm check-wholesale pricecheck_retail_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" style="<%= inactive_field %>" name="retail_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= FormatNumber(rs_getdetails.Fields.Item("price").Value,2)%>" data-column="price" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Retail price" data-pricecheck="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm check-wholesale pricecheck_wlsl_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" style="<%= inactive_field %>" name="wholesale_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= FormatNumber(rs_getdetails.Fields.Item("wlsl_price").Value,2)%>" data-column="wlsl_price" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Wholesale price" data-pricecheck="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td class="align-middle">
			<input class="form-control form-control-sm" style="<%= inactive_field %>" name="vendor-sku_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" <% if rs_getdetails.fields.item("detail_code").value <> "" then%>value="<%=(rs_getdetails.Fields.Item("detail_code").Value)%>"<% else %>value=" "<% end if %> data-column="detail_code" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Vendor SKU">
		</td>
		<td class="text-center align-middle">
			<% 	if (rs_getdetails.Fields.Item("active").Value) = 1 then
					var_checked = "checked"
				else
					var_checked = ""
				end if

			%>
			<div class="custom-control custom-checkbox">
				<input class="custom-control-input" name="active_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" id="active_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="checkbox" value="1" <%= var_checked %>  data-column="active" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-unchecked="0" data-friendly="Detail active">
				<label class="custom-control-label" for="active_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
					&nbsp;
				</label>
			  </div>
		</td>
		<td class="text-nowrap align-middle">
			<% if rs_getdetails.Fields.Item("DateLastPurchased").Value <> "" then %>				
				<span role="button" class="date_expand" id="last_sold_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-container="body" data-toggle="popover" data-placement="left" data-html="true" data-trigger="focus" data-content='Loading <i class="fa fa-spinner fa-spin ml-3"></i>' data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
					<%= FormatDateTime(rs_getdetails.Fields.Item("DateLastPurchased").Value,2)%>
				</span>
			<% end if %>
		</td>
		<td class="align-middle">
			<span class="input_move" name="input_move_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"><span><span class="btn btn-sm btn-secondary font-weight-bold copyid" name="copy_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-id="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"><i class="fa fa-copy-far fa-lg"></i></span>
			<span class="btn btn-sm btn-secondary font-weight-bold moveid" name="move_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-id="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">M</span>
		</td>
	</tr>
	<tr class="expanded-details <%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" style="display:none">
		<td class="pl-3 pr-3 pb-4" colspan="17">
			<table class="table table-sm table-borderless w-auto">
				<thead class="table-secondary">
					<tr>
						<th>Added</th>
						<th>Bin</th>
						<th>Free</th>
						<th>Free Qty</th>
						<th>Weight (Ounces)</th>
						<th>Wearable</th>
						<th>Materials</th>
						<th>Colors</th>
					</tr>
				</thead>
				<tr>
					<td>
						<% if rs_getdetails.Fields.Item("DateAdded").Value <> "" then %>
							<%= FormatDateTime(rs_getdetails.Fields.Item("DateAdded").Value, 2)%>
						<% end if %>
					</td>
					<td>
						<input class="form-control form-control-sm" style="width:50px" name="bin_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" id="bin_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%=(rs_getdetails.Fields.Item("BinNumber_Detail").Value)%>" data-column="BinNumber_Detail" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Bin #">
					</td>
					<td>
						<select class="form-control form-control-sm" name="free" data-column="free" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Free threshold">
							<option value="<% if Not isNull(rs_getdetails.Fields.Item("free").Value) then %><%=(rs_getdetails.Fields.Item("free").Value)%><% else %>0<% end if %>" selected>
							<% if (rs_getdetails.Fields.Item("free").Value) = 0 then %>
							Not free
							<% else %>
							<%=(rs_getdetails.Fields.Item("free").Value)%>
							<% end if%>
							</option>
							<option value="0">Not free</option>
							<option value="30">$30 +</option>
							<option value="50">$50 +</option>
							<option value="75">$75 +</option>
							<option value="100">$100 +</option>
							<option value="150">$150 +</option>
						</select>
					</td>
					<td>
						<input class="form-control form-control-sm" style="width:50px" name="free-qty_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" id="free-qty_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" <% if rs_getdetails.fields.item("Free_QTY").value <> "" then %>value="<%=(rs_getdetails.Fields.Item("Free_QTY").Value)%>"<% else %>value=" "<% end if %> data-column="Free_QTY" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" style="content: '1';" data-friendly="Free qty">  
					</td>
					<td>
						<input class="form-control form-control-sm" style="width:50px" name="weight_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" <% if rs_getdetails.fields.item("weight").value <> "" then %>value="<%=(rs_getdetails.Fields.Item("weight").Value)%>"<% else %>value=" "<% end if %> data-column="weight" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Weight">
					</td>
					<td>
						<select class="form-control form-control-sm w-auto select-detail-wearable-materials variant-wearable" name="wearable_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" id="wearable_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="wearable_material" data-friendly="Wearable" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
							<% if rs_getdetails.Fields.Item("wearable_material").Value <> "" then %>
								<option value="<%= rs_getdetails.Fields.Item("wearable_material").Value %>" selected><%= rs_getdetails.Fields.Item("wearable_material").Value %></option>  
	
							<%
							end if ' if jewelry is not null
							%>	
							<option>Select wearable material...</option>					
							<%rs_getWearableMaterials.MoveFirst()%>
							<% While NOT rs_getWearableMaterials.EOF %>
							<option <% if rs_getdetails("wearable_material") <> "" then %><%=IIF(Instr(rs_getWearableMaterials("material_name"), "__") > 0, " disabled", "")%><% end if %> value="<%= Trim(rs_getWearableMaterials("material_name")) %>"><%= Trim(rs_getWearableMaterials("material_name")) %></option>
							<% 
								rs_getWearableMaterials.MoveNext()
								Wend
							%> 
						</select>
					</td>
					<td style="width:300px">
						<select  class="select-detail-materials variant-materials" name="materials_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" id="materials_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="detail_materials" data-friendly="Materials" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"  multiple>
							<% if rs_getdetails.Fields.Item("detail_materials").Value <> "" and Instr(rs_getdetails.Fields.Item("detail_materials").Value, "null") = 0 then
								' break full text stored values out into an array to have an <option> selected for each entry

								selected_materials_array = getUniqueItems(split(rs_getdetails.Fields.Item("detail_materials").Value,","))
							end if ' if jewelry is not null
							%>	
							<option>Select materials...</option>					
							<%rs_getMaterials.MoveFirst()%>
							<% While NOT rs_getMaterials.EOF %>
								<option <% if rs_getdetails("detail_materials") <> "" then %><%=IIF(isItemInArray(selected_materials_array, rs_getMaterials("material_name")), "selected ","")%><% end if %> <%=IIF(Instr(rs_getMaterials("material_name"), "__") > 0, " disabled", "")%> value="<%= Trim(rs_getMaterials("material_name")) %>"><%= Trim(rs_getMaterials("material_name")) %></option>
								<%rs_getMaterials.MoveNext()%>
							<% Wend %>					

						</select>
					</td>															
					<td style="width:300px">
						<select class="variant-colors select-colors" name="colors_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" id="colors_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="colors" data-friendly="Colors" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"  multiple>
							<% if rs_getdetails.Fields.Item("colors").Value <> "" and Instr(rs_getdetails.Fields.Item("colors").Value, "null") = 0  then
							    ' break full text stored values out into an array to have an <option> selected for each entry
								jewelry_array = getUniqueItems(split(rs_getdetails.Fields.Item("colors").Value," "))

							end if ' if jewelry is not null
							%>	
							
							<option>Select colors...</option>					
						<% for each x in color_array %>
							<option <% if rs_getdetails("colors") <> "" then %><%=IIF(isItemInArray(jewelry_array, x), "selected ","")%><% end if %> value="<%= Trim(x) %>"><%= Trim(x) %></option>
						<% next %>
						</select>

					</td>
				</tr>
			</table>
				<div class="form-group w-50">
					<label class="font-weight-bold" for="detail-notes_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">Item notes</label> 
					<textarea class="form-control form-control-sm" name="detail-notes_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" id="detail-notes_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="detail_notes" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Detail notes" class="detail_notes"><%=(rs_getdetails.Fields.Item("detail_notes").Value)%></textarea>				
				</div>
		</td>
	</tr>
</tbody>
	<% 
	rs_getdetails.MoveNext()
'	Wend
	If rs_getdetails.EOF Then Exit For  ' ====== PAGING
	Next ' ====== PAGING
	end if ' if recordset is not empty
	%>
</table>
<div align="center">
	<div class="paging paging-div">
		<!--#include file="products/inc_paging.asp" -->
	</div>
</div>



<div class="modal fade" id="modal-combine" tabindex="-1" role="dialog">
	<div class="modal-dialog" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Combine another product into this product</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body small">
				This will transfer, details, images, jewelry & photo reviews, set old # to inactive, insert (PREPEND) to front of detail description, and automatically write comments to old #
				<div class="form-group">
					<label class="font-weight-bold" for="combine_productid" data-friendly="SEO title">Product # to move FROM</label>
					<input class="form-control form-control-sm" name="combine_productid" id="combine_productid" type="text">
				</div>
				<div class="form-group">
					<label class="font-weight-bold" for="combine_detailinfo" data-friendly="SEO title">Color or detail info</label>
					<input class="form-control form-control-sm" name="combine_detailinfo" id="combine_detailinfo" type="text">
				</div>				
				<button id="combine_now" class="btn btn-sm btn-primary">Combine now</button>
		</div>
	  </div>
	</div>
  </div>


<!-- Begin Tarrif modal -->
<div class="modal fade" id="modal-tariff-info" tabindex="-1" role="dialog"  aria-labelledby="modal-tariff-info" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Tariff Information</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
			<h6>7117.90.9000 - Any base material jewelry</h6>
			Any items made from steel, brass, bronze, copper, glass, acrylic, wood, horn, or bone
			<h6 class="mt-3">7113.11.5000 - Base material from solid silver</h6>
			Any items where the base is solid silver. This also includes silver that has been plated with gold or platinum. It does not matter if it has stones in it. All that matters is what the base metal is made from.
			<h6 class="mt-3">7116.20.0500 - Semiprecious stone plugs (Natural, synthetic, or reconstructed)</h6>
			Any stone plug that we have that can naturally be found (So no cat eye, goldstone, or glass)
		</div>
	  </div>
	</div>
</div>
<!-- End Tariff Modal  -->


<!-- BEGIN EDITS LOG MODAL WINDOW -->
<div class="modal fade" id="modal-edits-log" tabindex="-1" role="dialog"  aria-labelledby="modal-edits-log" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Edits & Scan Log</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div id="show-edits" class="modal-body">
			
		</div>
	  </div>
	</div>
</div>
<!-- END EDITS LOG MODAL WINDOW -->

<!-- BEGIN MANAGE TAGS MODAL WINDOW -->
<div class="modal fade" id="modal-show-tags" tabindex="-1" role="dialog"  aria-labelledby="modal-show-tags" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Manage Tags</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div id="show-tags" class="modal-body">
			
		</div>
	  </div>
	</div>
</div>
<!-- END MANAGE TAGS MODAL WINDOW -->

<!-- BEGIN MANAGE MATERIALS MODAL WINDOW -->
<div class="modal fade" id="modal-show-materials" tabindex="-1" role="dialog"  aria-labelledby="modal-show-materials" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Manage Materials</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div id="show-materials" class="modal-body">
			
		</div>
	  </div>
	</div>
</div>
<!-- END MANAGE MATERIALS MODAL WINDOW -->

<!-- BEGIN MANAGE CATEGORIES MODAL WINDOW -->
<div class="modal fade" id="modal-show-categories" tabindex="-1" role="dialog"  aria-labelledby="modal-show-categories" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Manage Categories</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div id="show-categories" class="modal-body">
			
		</div>
	  </div>
	</div>
</div>
<!-- END MANAGE CATEGORIES MODAL WINDOW -->

<% end if ' if product is found then display page %>
<div style="height: 1000px;"></div>
<div id="enlarge_footer_image" class="fixed-bottom enlarge_footer_image" style="display:none;height: 240px;width: 300px;padding: 10px;bottom: 70px;left: 0;right: 0;"></div>
<div class="fixed-bottom form-inline small py-1" style="background: rgba(25, 25, 25, .8)">
	<div class="container-fluid">
		<div class="row">
		  <div class="col col-10 ajax-update">
			<span class="font-weight-bold text-light" id="detail_images"></span><span id="sort-message"></span>
			<input class="form-control form-control-sm ml-3 mr-2" style="display:none" type="text" placeholder="Image description" value="" data-imgid="" id="input-img-description"><span class="img-thumb-clone"></span>
		  </div>
		  <div class="col col-2 text-right">
			<button class="btn btn-sm btn-danger px-2" type="button" id="edit_images_link"><i class="fa fa-trash-alt"></i></button>
		  </div>
		</div>
	  </div>

</div>
</body>
<script type="text/javascript" src="/js/popper.min.js"></script>
<script type="text/javascript" src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui.min.js"></script>
<script type="text/javascript" src="/js/bootstrap-v4.min.js"></script>
<script type="text/javascript" src="/js/redactor.js"></script>
<script type="text/javascript" src="/js/redactor-plugin-source.js"></script>
<script type="text/javascript" src="/js/chosen/chosen.jquery.js"></script>
<script type="text/javascript" src="scripts/dropzone.js"></script>
<script type="text/javascript" src="scripts/jquery.validate.min.js"></script>
<script type="text/javascript" src="scripts/product-edit-version2.js?v=080921"></script>

</html>
<%
Set rs_getuser = Nothing
Set rs_getdetails = Nothing
Set rs_getproduct = Nothing
Set rs_getbrand = Nothing
Set rs_getTags = Nothing
Set rsGetInactive = Nothing
Set rsGetGauges = Nothing
set rs_getsections = nothing
set rs_getTags = nothing
set rs_getWearableMaterials = nothing
set rs_getMaterials = nothing
set rs_getCategories = nothing
DataConn.Close()

Function getUniqueItems(arrItems)
	Dim objDict, strItem

	Set objDict = Server.CreateObject("Scripting.Dictionary")

	For Each strItem in arrItems
		objDict.Item(strItem) = 1
	Next

	getUniqueItems = objDict.Keys
End Function

Function isItemInArray(theArray, theValue)
    dim i, fnd
    fnd = False
    For i = 0 to UBound(theArray)
        If Trim(theArray(i)) = Trim(theValue) Then
            fnd = True
            Exit For
        End If
    Next
    isItemInArray = fnd
End Function
%>