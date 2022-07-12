<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


if request.querystring("remove-filters") = "yes" then
    session("filter-category") = ""
end if

if request.querystring("filter-category") <> "" then
    session("filter-category") = request.querystring("filter-category")
end if

if session("filter-category") <> "" then
    sql = "WHERE category = ?"
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.Prepared = true
objCmd.CommandText = "SELECT * FROM tbl_sitemap_searches " & sql & " ORDER BY id ASC"
    if session("filter-category") <> "" then
        objCmd.Parameters.Append(objCmd.CreateParameter("filter",200,1,30,session("filter-category")))
    end if
set rsGetSEO = Server.CreateObject("ADODB.Recordset")
rsGetSEO.CursorLocation = 3 'adUseClient
rsGetSEO.Open objCmd
rsGetSEO.PageSize = 25
total_records = rsGetSEO.RecordCount
intPageCount = rsGetSEO.PageCount

' Variables for paging
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
%>
<html>
<head>
<title>SEO Product Search Manager</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>SEO Product Search Manager</h5>

<form class="form-inline" id="frm-filters" formaction="seo-product-search-manager.asp">
<select class="form-control form-control-sm" name="filter-category" id="filter-category">
        <option value="">Filter by Category</option>
	<option value="aftercare">Aftercare</option>
	<option value="barbell">Barbell</option>
    <option value="brand">By Brand</option>
    <option value="gauge">By Gauge</option>
    <option value="keyword">By Keyword</option>
	<option value="captive">Captive</option>
	<option value="circular">Circular Barbell</option>
	<option value="gear">BAF gear</option>
	<option value="curved">Curved</option>
	<option value="hanging">Hanging styles</option>
	<option value="labret">Labret</option>
	<option value="ends">Loose ends</option>
	<option value="navel">Navel jewelry</option>
	<option value="nipple">Nipple Jewelry</option>
	<option value="nose">Nose</option>
	<option value="plugs">Plugs</option> 
	<option value="regular">Regular jewelry</option>
	<option value="retainer">Retainer</option>
	<option value="septum">Septum</option>
	<option value="tapers">Tapers</option>
	<option value="twists">Twists</option>
	<option value="weight">Weights</option>
    </select>
    </form>
<%
if session("filter-category") <> "" then
%>
        <a class="text-danger my-2 d-block" href="seo-product-search-manager.asp?remove-filters=yes"><i class="fa fa-times-circle pr-2"></i>Remove filters</a>
<%
end if
%>

</div>
<!--#include file="seo/seo-paging.asp"-->
<table class="table table-striped table-hover order-list ajax-update">
<thead class="thead-dark">
<tr>
    <td colspan="5">
        <button class="btn btn-sm btn-primary" type="button" id="add-row">Add Row</button></td>
</tr>
  <tr>
    <th class="sticky-top">URL</th>
    <th class="sticky-top">Category / Hero Image</th>
    <th class="sticky-top">Title (On Page)</th>
	<th class="sticky-top">SEO Title</th>
	<th class="sticky-top">SEO Description</th>
  </tr>
</thead>
<tbody>
<% 
if NOT rsGetSEO.eof then
rsGetSEO.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetSEO.PageSize 

%>
    <tr class="<%= rsGetSEO.Fields.Item("id").Value %>">
      <td><textarea class="form-control mb-1" rows="4" name="url_<%= rsGetSEO.Fields.Item("id").Value %>" data-column="url" data-id="<%= rsGetSEO.Fields.Item("id").Value %>"><%= rsGetSEO.Fields.Item("url").Value %></textarea>
        <div class="my-2">
                <i class="fa fa-trash-alt text-danger pointer" data-id="<%= rsGetSEO.Fields.Item("id").Value %>"></i>
            <a class="mx-5 text-dark" href="../products.asp?<%= rsGetSEO.Fields.Item("url").Value %>" target="_blank"><i class="fa fa-external-link"></i> Open URL</a>
            <a class="text-dark" href="../products.asp?<%= Server.HTMLEncode(rsGetSEO.Fields.Item("canonical_url").Value & "") %>" target="_blank"><i class="fa fa-external-link"></i> Open Canonical</a>
     </div>
        <input class="form-control" type="text" value="<%= Server.HTMLEncode(rsGetSEO.Fields.Item("canonical_url").Value & "") %>" placeholder="Canonical URL : /products.asp?" name="canonical_url__<%= rsGetSEO.Fields.Item("id").Value %>" data-column="canonical_url" data-id="<%= rsGetSEO.Fields.Item("id").Value %>">
      <td>   
        <select class="form-control" id="filter-cagegory" name="category_<%= rsGetSEO.Fields.Item("id").Value %>" data-column="category" data-id="<%= rsGetSEO.Fields.Item("id").Value %>">
                <option selected value="<%= rsGetSEO.Fields.Item("category").Value %>"><%= rsGetSEO.Fields.Item("category").Value %></option>
                <option value="">Filtering category ----</option>
                <option value="aftercare">Aftercare</option>
                <option value="barbell">Barbell</option>
                <option value="brand">By Brand</option>
                <option value="gauge">By Gauge</option>
                <option value="keyword">By Keyword</option>
                <option value="captive">Captive</option>
                <option value="circular">Circular Barbell</option>
                <option value="gear">BAF gear</option>
                <option value="curved">Curved</option>
                <option value="hanging">Hanging styles</option>
                <option value="labret">Labret</option>
                <option value="ends">Loose ends</option>
                <option value="navel">Navel jewelry</option>
                <option value="nipple">Nipple Jewelry</option>
                <option value="nose">Nose</option>
                <option value="plugs">Plugs</option> 
                <option value="regular">Regular jewelry</option>
                <option value="retainer">Retainer</option>
                <option value="septum">Septum</option>
                <option value="tapers">Tapers</option>
                <option value="twists">Twists</option>
                <option value="weight">Weights</option>
                </select>
            <input class="form-control mt-2" type="text" name="hero_image_<%= rsGetSEO.Fields.Item("id").Value %>" value="<%= rsGetSEO.Fields.Item("hero_image").Value %>"  data-column="hero_image" data-id="<%= rsGetSEO.Fields.Item("id").Value %>" placeholder="Hero image name">
            <textarea class="form-control mt-2" rows="3" name="extra_keywords_<%= rsGetSEO.Fields.Item("id").Value %>" placeholder="URL Keywords" data-column="extra_keywords" data-id="<%= rsGetSEO.Fields.Item("id").Value %>"><%= rsGetSEO.Fields.Item("extra_keywords").Value %></textarea>
            </td>
                <td><textarea class="form-control meta_title_onpage" rows="4" type="text" maxlength="100" name="meta_title_onpage_<%= rsGetSEO.Fields.Item("id").Value %>" data-column="meta_title_onpage" data-id="<%= rsGetSEO.Fields.Item("id").Value %>"><%= rsGetSEO.Fields.Item("meta_title_onpage").Value %></textarea>
                    <div id="remaining_title_onpage_<%= rsGetSEO.Fields.Item("id").Value %>"></div></td>
      <td><textarea class="form-control meta-title" rows="4" type="text" maxlength="60" name="meta_title_<%= rsGetSEO.Fields.Item("id").Value %>" data-column="meta_title" data-id="<%= rsGetSEO.Fields.Item("id").Value %>"><%= rsGetSEO.Fields.Item("meta_title").Value %></textarea>
        <div id="remaining_title_<%= rsGetSEO.Fields.Item("id").Value %>"></div></td>
      <td><textarea class="form-control meta-description" rows="4" type="text" maxlength="200" name="meta_description_<%= rsGetSEO.Fields.Item("id").Value %>" data-column="meta_description" data-id="<%= rsGetSEO.Fields.Item("id").Value %>"><%= rsGetSEO.Fields.Item("meta_description").Value %></textarea>
        <div id="remaining_desc_<%= rsGetSEO.Fields.Item("id").Value %>"></div></td>
    </tr>
    <% 
rsGetSEO.MoveNext()
If rsGetSEO.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
end if 'NOT rsGetSEO.eof
%>
</tbody>
</table>

<!--#include file="seo/seo-paging.asp"-->
</body>
</html>
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript">

	//url to to do auto updating
	var auto_url = "seo/ajax-update-seo-product-search.asp"
</script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
	auto_update();

    	$('#add-row').click(function () {
		$.ajax({
		method: "post",
		dataType: "json",
		url: "seo/ajax-add-row.asp"
		})
		.done(function(json, msg) {
            if (json.status == 'success') {
				console.log("row added");
                var newRow = $("<tr>");
                var cols = "";

            cols += '<td><textarea class="form-control" rows="4" name="url_' + json.id + '" data-column="url" data-id="' + json.id + '"></textarea></td>';
            cols += '<td><select class="form-control" id="filter-cagegory" name="category_' + json.id + '" data-column="category" data-id="' + json.id + '"><option value="aftercare">Aftercare</option><option value="barbell">Barbell</option><option value="brand">By Brand</option><option value="gauge">By Gauge</option><option value="captive">Captive</option><option value="circular">Circular Barbell</option><option value="gear">BAF gear</option><option value="curved">Curved</option><option value="hanging">Hanging styles</option><option value="labret">Labret</option><option value="ends">Loose ends</option><option value="navel">Navel jewelry</option><option value="nipple">Nipple Jewelry</option><option value="nose">Nose</option><option value="plugs">Plugs</option> <option value="regular">Regular jewelry</option><option value="retainer">Retainer</option><option value="septum">Septum</option><option value="tapers">Tapers</option><option value="twists">Twists</option><option value="weight">Weights</option></select><textarea class="form-control mt-2" rows="3" name="extra_keywords_' + json.id + '" placeholder="URL Keywords" data-column="extra_keywords" data-id="' + json.id + '"></textarea></td>';
            cols += '<td><textarea class="form-control meta_title_onpage" rows="4" type="text" maxlength="60" name="meta_title_onpage_' + json.id + '" data-column="meta_title_onpage" data-id="' + json.id + '"></textarea><div id="remaining_title_onpage_' + json.id + '"></div></td>';
            cols += '<td><textarea class="form-control meta-title" rows="4" type="text" maxlength="60" name="meta_title_' + json.id + '" data-column="meta_title" data-id="' + json.id + '"></textarea><div id="remaining_title_' + json.id + '"></div></td>';

            cols += '<td><textarea class="form-control meta-description" rows="4" type="text" maxlength="200" name="meta_description_' + json.id + '" data-column="meta_description" data-id="' + json.id + '"></textarea><div id="remaining_desc_' + json.id + '"></div></td>';
            newRow.append(cols);
            $("table.order-list").prepend(newRow);
			}
			if (json.status == 'fail') {
				console.log("row not added");
			}
		})
		.fail(function(json, msg) {
			console.log("site error");
		})
	});  // END redeem points

       $(document).on("keypress", '.meta_title_onpage', function() {
        id = $(this).attr('data-id');

        if(this.value.length > 50){
            return false;
        }
        $("#remaining_title_onpage_" + id).html("Remaining characters : " +(50 - this.value.length));
        });

       $(document).on("keypress", '.meta-title', function() {
        id = $(this).attr('data-id');

        if(this.value.length > 50){
            return false;
        }
        $("#remaining_title_" + id).html("Remaining characters : " +(50 - this.value.length));
        });
       
        $(document).on("keypress", '.meta-description', function() {
        id = $(this).attr('data-id');

        if(this.value.length > 155){
            return false;
        }
        $("#remaining_desc_" + id).html("Remaining characters : " +(155 - this.value.length));
        });

    $(document).on("change", '#filter-category', function() {
        $('#frm-filters').submit();
    });

    // Delete row
    $(document).on("click", '.fa-trash-alt', function() {
        id = $(this).attr('data-id');
        
        $.ajax({
		method: "post",
        url: "seo/ajax-delete-search-row.asp",
        data: {id:id}
		})
		.done(function(msg) {
            $('.' + id).fadeOut('slow');
		})
		.fail(function(msg) {
			console.log("site error");
		})
    });

    

</script>
<%
rsGetSEO.Close()
Set rsGetSEO = Nothing
%>
