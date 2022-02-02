<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

if session("filterby-time") <> "" then
    sql_filter_time = session("filterby-time")
else
    sql_filter_time = 6
end if

if session("filterby-brand") <> "" then
    sql_filter_brand = " AND brandname = ? "
else
    sql_filter_brand = ""
end if


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.Prepared = true
objCmd.CommandText = "SELECT j.ProductID, j.title, j.description, j.picture, j.date_added, j.salediscount, CAST(j.ProductNotes AS VARCHAR(MAX)) as ProductNotes, j.type, j.brandname, MAX(pd.DateLastPurchased) as 'LastPurchaseDate', MIN(pd.DateLastPurchased) as 'OldestPurchaseDate' FROM jewelry j inner join dbo.ProductDetails pd on j.ProductID = pd.ProductID WHERE j.Active = 1 and j.jewelry not like '%save%' and j.title not like '%custom order%'  GROUP BY j.ProductID, j.title, j.description, j.picture, j.date_added, j.salediscount, CAST(j.ProductNotes AS VARCHAR(MAX)), j.type, j.brandname HAVING (MAX(ISNULL(pd.DateLastPurchased,'2000-01-01')) <= DATEADD(month, -" & sql_filter_time & ", j.date_added )) " & sql_filter_brand & " ORDER BY MAX(pd.DateLastPurchased)"

if session("filterby-brand") <> "" then
    objCmd.Parameters.Append(objCmd.CreateParameter("brand",8,1,50, session("filterby-brand") ))
end if

set rsGetProducts = Server.CreateObject("ADODB.Recordset")
rsGetProducts.CursorLocation = 3 'adUseClient
rsGetProducts.Open objCmd
rsGetProducts.PageSize = 50
total_records = rsGetProducts.RecordCount
intPageCount = rsGetProducts.PageCount

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

'===== GET BRANDS FOR SELECT MENU filtering
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT companyID, name FROM TBL_Companies WHERE display_AddEdit = 'yes' AND type = 'jewelry' ORDER BY name ASC"
Set rs_getbrand = objCmd.Execute()
%>
<!DOCTYPE html>
<html>
<head>
<title>Manage inventory not selling</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="mt-2 ml-2">
<h5>Manage inventory not selling</h5>

	<div class="form-group form-inline">
		<select class="ml-3 form-control form-control-sm" id="filter-bytime">
            <% if session("filterby-time") <> "" then %>
            <option value="<%= session("filterby-time") %>" selected="selected">More than <%= session("filterby-time") %> months
            </option>
            <% else %>
            <option value="6" selected="selected">Filter by last sold date</option>
            <% end if %>
		  <option value="6">Over 6 months from date added (default)</option>
		  <option value="9">Over 9 months from date added</option>
		  <option value="12">Over 1 year from date added</option>
		  <option value="24">Over 2 years from date added</option>
		  <option value="36">Over 3 years from date added</option>
		</select>

        <select class="form-control form-control-sm ml-3" id="filter-bybrand">
            <% if session("filterby-brand") <> "" then %>
            <option value="<%= session("filterby-brand") %>" selected="selected"><%= session("filterby-brand") %>
            </option>
            <% else %>
            <option value="" selected="selected">Filter by brand</option>
            <% end if %>
            <option value="">NO BRAND</option>
            <% 
            While NOT rs_getbrand.EOF 
            %>
            <option value="<%=(rs_getbrand.Fields.Item("name").Value)%>"><%=(rs_getbrand.Fields.Item("name").Value)%>
            </option>
            <% 
            rs_getbrand.MoveNext()
            Wend
            %>                
        </select>

        <button class="ml-3 btn btn-sm btn-secondary" id="submit-filters">Update filters</button>
	</div>


</div>
<!--#include file="inventory/inc-old-products-paging.asp" -->

			<table class="table table-striped table-hover table-sm table-bordered">
				<thead class="thead-dark">
				  <tr>
					<th scope="col">Product</th>
					<th scope="col">Purchase Info</th>
					<th scope="col">Status</th>
					<th scope="col">Sale</th>
				  </tr>
				</thead>
				<tbody>
<% 
if NOT rsGetProducts.EOF then
rsGetProducts.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetProducts.PageSize 

if rsGetProducts.Fields.Item("salediscount").Value > 0 then
    sale_class = "table-success"
else
    sale_class = "table-danger"
end if
%>
            <tr class="<%= sale_class %>" id="rowid-<%= rsGetProducts.Fields.Item("ProductID").Value %>">
                <td scope="row"  style="width: 50%">
                    <div class="image-container float-left">
                        <a class="font-weight-bold" href="product-edit.asp?ProductID=<%= rsGetProducts.Fields.Item("ProductID").Value %>" target="_blank"><img class="mr-2" src="https://bodyartforms-products.bodyartforms.com/<%= rsGetProducts.Fields.Item("picture").Value %>" /></a>
                    </div>
                    
                    <div>
                        <div class="font-weight-bold small"><%= rsGetProducts.Fields.Item("ProductNotes").Value %></div>
                        <div id="variants-<%= rsGetProducts.Fields.Item("ProductID").Value %>">
                            <button class="btn btn-sm btn-outline-secondary show-variants" data-id="<%= rsGetProducts.Fields.Item("ProductID").Value %>">View variants</button>
                        </div>
                    </div>
                    
                </td>
                <td>
                    <div class="font-weight-bold small"><%= rsGetProducts.Fields.Item("brandname").Value %></div>
                    Date added: <%= rsGetProducts.Fields.Item("date_added").Value %>
                    <div class="my-1">Last purchase: <%= rsGetProducts.Fields.Item("LastPurchaseDate").Value %></div>
                    Oldest purchase: <%= rsGetProducts.Fields.Item("OldestPurchaseDate").Value %>
                </td>
                <td>
                    <div class="form-inline">
                    <select class="form-control form-control-sm field-update" data-id="<%= rsGetProducts.Fields.Item("ProductID").Value %>" data-column="type">
                        <option value="<%= rsGetProducts.Fields.Item("type").Value %>"  selected="selected"><%= rsGetProducts.Fields.Item("type").Value %></option>
                        <option value="limited">limited</option>
                        <option value="clearance">clearance</option>
                        <option value="One time buy">One time buy</option>
                        <option value="Discontinued">Discontinued</option>
                        <option value="None">None</option>
                    </select>
                </div>
                   
                </td>
                <td>
                    <div class="form-inline">
                    <select class="form-control form-control-sm field-update" data-id="<%= rsGetProducts.Fields.Item("ProductID").Value %>" data-column="saleDiscount">
                        <option value="<%= rsGetProducts.Fields.Item("saleDiscount").Value %>"  selected="selected"><%= rsGetProducts.Fields.Item("saleDiscount").Value %>%</option>
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
                
            </tr>
            <% 
rsGetProducts.MoveNext()
If rsGetProducts.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING

end if ' if NOT rsGetProducts.EOF
%>
				</tbody>
			  </table>


<!--#include file="inventory/inc-old-products-paging.asp" -->

<script type="text/javascript">
// Update database based on which select fields are changed
$(document).on("change", '.field-update', function() { 
        var value = $(this).val();
        var productid = $(this).attr("data-id");
        var column = $(this).attr("data-column");
		
		$.ajax({
		method: "POST",
        context: this,
		url: "inventory/inc-old-products-updates.asp",
		data: {productid: productid, column: column, value: value}
		})
		.done(function( msg ) {
            $(this).after('<i class="fa fa-lg fa-check text-success ml-2"></i>');

            // Only change row color if a sale is being activated
            if (column == 'saleDiscount') {
                $('#rowid-' + productid).removeClass("table-danger");
                $('#rowid-' + productid).addClass("table-success");
            }
		})
		.fail(function(msg) {
            $(this).after('<i class="fa fa-lg fa-times-circle text-danger ml-2"></i> Update failed');
		});
	});

    // Reload page on filter changes
    $(document).on("click", '#submit-filters', function() { 
        var filter_time = $('#filter-bytime').val();
        var filter_brand = $('#filter-bybrand').val();
		
		$.ajax({
		method: "POST",
		url: "inventory/inc-old-products-sessions.asp",
		data: {filter_time: filter_time, filter_brand: filter_brand}
		})
		.done(function( msg ) {
            window.location = "inventory-manage-old-products.asp";
		})
		.fail(function(msg) {
            alert("Setting filter session failed");
		});
	});
    
    // Load variants once show variant button is clicked
    $(document).on("click", '.show-variants', function() { 
        var productid = $(this).attr("data-id");

		$('#variants-' + productid).load("inventory/inc-old-products-variants.asp", {productid: productid}, function() {       });	
	});	
</script>
</body>
</html>