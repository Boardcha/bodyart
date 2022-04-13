<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%	
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.Prepared = true
objCmd.CommandText = "SELECT j.productid, j.brandname, j.active, fp.min_gauge, fp.max_gauge, j.picture, j.picture_400, j.title, j.seo_meta_title, j.seo_meta_description, color_tags, j.material FROM FlatProducts AS fp INNER JOIN jewelry AS j ON fp.productid = j.ProductID WHERE j.active = 1 and (seo_meta_title IS NULL OR seo_meta_description IS NULL)  ORDER BY j.ProductID DESC"
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
<title>SEO Title & Description Tagging</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>SEO Title & Description Tagging</h5>

</div>
<!--#include file="seo/seo-paging.asp"-->
<table class="table table-striped table-hover order-list">
<thead class="thead-dark">
  <tr>
    <th class="sticky-top">Information</th>
    <th class="sticky-top">Titles</th>
	<th class="sticky-top">SEO Description</th>
  </tr>
</thead>
<tbody class="ajax-update">
<% 
if NOT rsGetSEO.eof then
rsGetSEO.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetSEO.PageSize 

%>
    <tr>
      <td style="width:25%">
        <a class="text-dark" href="product-edit.asp?ProductID=<%= rsGetSEO.Fields.Item("productid").Value %>" target="_blank"><img class="pull-left mr-2" style="width:120px" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetSEO.Fields.Item("picture_400").Value %>" alt="<%= rsGetSEO.Fields.Item("title").Value %>" />
        </a>
        <%= rsGetSEO.Fields.Item("min_gauge").Value %>
        <% if rsGetSEO.Fields.Item("max_gauge").Value <> rsGetSEO.Fields.Item("min_gauge").Value then %>
        thru <%= rsGetSEO.Fields.Item("max_gauge").Value %>
        <% end if %><br/>
        <%= rsGetSEO.Fields.Item("brandname").Value %><br/>
        <%= rsGetSEO.Fields.Item("material").Value %>
    </td>
    <td style="width:35%">
        <div class="mb-2"><%= rsGetSEO.Fields.Item("title").Value %></div>
        SEO Title
        <textarea class="form-control seo_title" rows="3" type="text" maxlength="100" name="seo_meta_title_<%= rsGetSEO.Fields.Item("productid").Value %>" id="seo_title_<%= rsGetSEO.Fields.Item("productid").Value %>" data-column="seo_meta_title" data-id="<%= rsGetSEO.Fields.Item("productid").Value %>"><%= rsGetSEO.Fields.Item("seo_meta_title").Value %></textarea>
        <div id="remaining_title_<%= rsGetSEO.Fields.Item("productid").Value %>"></div>
        <div id="msg-seo-title-<%= rsGetSEO.Fields.Item("productid").Value %>"></div>
    </td>
      <td style="width:40%"><textarea class="form-control meta-description" rows="4" type="text" maxlength="200" name="seo_meta_description_<%= rsGetSEO.Fields.Item("productid").Value %>" data-column="seo_meta_description" data-id="<%= rsGetSEO.Fields.Item("productid").Value %>"><%= rsGetSEO.Fields.Item("seo_meta_description").Value %></textarea>
        <div id="remaining_desc_<%= rsGetSEO.Fields.Item("productid").Value %>"></div></td>
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
	var auto_url = "seo/ajax-update-title-description-tagging.asp"
</script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
	auto_update();

       $(document).on("keypress", '.seo_title', function() {
        id = $(this).attr('data-id');

        if(this.value.length > 60){
            return false;
        }
        $("#remaining_title_" + id).html("Remaining characters : " +(60 - this.value.length));
        });
       
        $(document).on("keypress", '.meta-description', function() {
        id = $(this).attr('data-id');

        if(this.value.length > 160){
            return false;
        }
        $("#remaining_desc_" + id).html("Remaining characters : " +(160 - this.value.length));
        });

        // Duplicate title into SEO title field and then check to see if it's a duplicate
        $(".seo_title").change(function () {
            var productid = $(this).attr('data-id');
            var seo_title = $("#seo_title_" + productid).val();
            
            $.ajax({
            method: "POST",
            dataType: "json",
            url: "products/ajax-check-duplicate-title.asp",
            data: {seo_title: seo_title, id: productid}
            })
            .done(function( json, msg ) {
                if (json.status === "fail") {
                    $("#msg-seo-title-" + productid).html('Duplicate title found... Needs to be updated')
                    $("#msg-seo-title-" + productid).addClass("notice-red");
                    console.log('Duplicate found');
                } else {
                    $("#msg-seo-title-" + productid).html('')
                    $("#msg-seo-title-" + productid).removeClass("notice-red");
                    console.log('No duplicate found');
                }
            })
            .fail(function(msg) {
                $("#msg-seo-title-" + productid).html('Error checking duplicate')
                $("#msg-seo-title-" + productid).addClass("notice-red");
            });

        });	
</script>
<%
rsGetSEO.Close()
Set rsGetSEO = Nothing
%>
