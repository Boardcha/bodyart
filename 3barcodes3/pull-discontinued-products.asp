<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if rsGetUser.bof AND rsGetUser.eof then
    response.redirect "login.asp"
end if 



set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM jewelry WHERE to_be_pulled = 1 AND pull_completed = 0"
set rsGetProducts = Server.CreateObject("ADODB.Recordset")
rsGetProducts.CursorLocation = 3 'adUseClient
rsGetProducts.Open objCmd
rsGetProducts.PageSize = 1
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

%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
<title>Barcode management</title>
<link href="/CSS/baf.min.css?v=092519" rel="stylesheet" type="text/css" />
<script src="https://use.fortawesome.com/dc98f184.js"></script>
</head>

<body>
 <!--#include file="includes/scanners-header.asp" -->
<div class="p-3">
    <h6><%= total_records %> products to be pulled</h6>
    <% if not rsGetProducts.eof then %>
    <!--#include file="includes/inc-pull-products-paging.asp" -->
    <% rsGetProducts.AbsolutePage = intPage '======== PAGING
    For intRecord = 1 To rsGetProducts.PageSize 

        ' --- pull details
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "SELECT ID_Description, BinNumber_Detail, location, ProductDetailID, ProductDetails.ProductID as ProductID, qty, type, qty_counted_discontinued, item_pulled, Gauge, Length, ProductDetail1  FROM ProductDetails INNER JOIN TBL_GaugeOrder ON COALESCE (ProductDetails.Gauge, '') = COALESCE (TBL_GaugeOrder.GaugeShow, '') INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE (BinNumber_Detail = 37 OR BinNumber_Detail = 0) AND ProductDetails.ProductID = ? and ProductDetails.active = 1 ORDER BY ID_BarcodeOrder ASC, BinNumber_Detail ASC, location ASC"
        objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,15,rsGetProducts.Fields.Item("ProductID").Value ))
        'objCmd.Parameters.Append(objCmd.CreateParameter("who",200,1,50, rsGetUser.Fields.Item("name").Value ))

        set rsGetDetails = Server.CreateObject("ADODB.Recordset")
        rsGetDetails.CursorLocation = 3 'adUseClient
        rsGetDetails.Open objCmd

        if rsGetDetails.eof then
            details_message = "<div class='alert alert-danger'>No active items to be pulled</div>"
        end if

        ' ---- Check to see if there are any details left and if not, then send back json response
        ' --- pull details
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "SELECT item_pulled FROM ProductDetails WHERE ProductID = ? AND item_pulled = 0 AND active = 1 AND BinNumber_Detail = 0"
        objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, rsGetProducts("ProductID")  ))
        set rsCheckCompletion = objCmd.Execute()

        if NOT rsCheckCompletion.eof then
            display_scanner_field = "style=""display:none"""
        else
            display_scanner_field = " "
        end if
    %>
    <div style="display:none" id="message"></div>
    <div class="form-inline mb-3" id="frm-scan-to-bin" <%= display_scanner_field %>>
<label for="bin_number"></label>
            <input type="text" class="form-control" name="bin_number" id="bin_number" placeholder="Scan bin # to assign all">
            <button class="btn btn-primary ml-3" id="btn-finalize" id="btn-finalize" name="button" data-productid="<%= rsGetProducts.Fields.Item("ProductID").Value %>">FINALIZE</button>
    </div>
    <button class="btn btn-sm btn-warning mb-2" id="btn-delete" data-productid="<%= rsGetProducts.Fields.Item("ProductID").Value %>" type="button"><i class="fa fa-times-circle fa-lg"></i> Hide product</button>
    <table class="table table-bordered table-sm small">
            <thead class="thead-light">
                    <tr>
                      <th colspan="3">
                            <a href="/productdetails.asp?ProductID=<%= rsGetProducts.Fields.Item("ProductID").Value %>" target="_blank"><img src="https://bafthumbs-400.bodyartforms.com/<%= rsGetProducts.Fields.Item("picture").Value %>" style="width:50px;height:auto"></a>  
                        <%= rsGetProducts.Fields.Item("title").Value %>
                        <% if ISNULL(rsGetProducts.Fields.Item("who_pulled").Value) then %>
                        <button class="btn btn-primary btn-sm ml-4" id="btn-assign" data-productid="<%= rsGetProducts.Fields.Item("ProductID").Value %>">Pull Items</button>
                        <% else %>
                        <span class="badge badge-info ml-2" style="font-size:1.25em">Assigned to: <%= rsGetProducts.Fields.Item("who_pulled").Value %></span>
                        <% end if %>
                        
                    </th>
                    </tr>
                    <tr>
                            <th width="20%">
                              Location
                          </th>
                          <th width="30%">
                                Qty
                            </th>
                            <th width="50%">
                                    Item
                                </th>
                          </tr>
            </thead>
    <tbody>
    <% while not rsGetDetails.eof 
    if rsGetDetails.Fields.Item("item_pulled").Value = 1 then
        row_success = " table-success "
        checkmark_success = " text-success "
        disabled_input = " disabled "
    else
        row_success = " "    
        checkmark_success = " text-black-50 "
        disabled_input = " "
    end if

    if rsGetDetails.Fields.Item("qty_counted_discontinued").Value = -1 then
        var_qty = rsGetDetails.Fields.Item("qty").Value
        qty_description = ""
    else
        var_qty = rsGetDetails.Fields.Item("qty_counted_discontinued").Value
        qty_description = "Your count"
    end if
    %>
    <tr class="tr-<%= rsGetDetails.Fields.Item("ProductDetailID").Value %> <%= row_success %>"><td>
            <%= rsGetDetails.Fields.Item("ID_Description").Value %>&nbsp;
			<% if rsGetDetails.Fields.Item("BinNumber_Detail").Value <> 0 then %>
            BIN <%= rsGetDetails.Fields.Item("BinNumber_Detail").Value %>
            <% end if %>
            <span class="ml-1"><%= rsGetDetails.Fields.Item("location").Value %></span>
    </td>
    <td>
                        <span class="input-group-sm form-inline">
                            <i class="fa fa-check-circle fa-2x  mr-4 <%= checkmark_success %> confirm-item-pulled check-<%= rsGetDetails.Fields.Item("ProductDetailID").Value %>" data-detailid="<%= rsGetDetails.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetDetails.Fields.Item("ProductID").Value %>"></i><input type="text" class="form-control w-25" name="qty_counted" id="qty_<%= rsGetDetails.Fields.Item("ProductDetailID").Value %>" value=" <%= var_qty %>" <%= disabled_input %>>
                            <span class="ml-1"><%= qty_description %></span>
                        </span>
    </td>
<td>
        <% If (rsGetDetails.Fields.Item("Gauge").Value) <> "" Then %>
        <%= Server.HtmlEncode(rsGetDetails.Fields.Item("Gauge").Value)%>
    <% end if %>
    &nbsp;&nbsp;
    <% If (rsGetDetails.Fields.Item("Length").Value) <> "" Then %>
        <%= Server.HtmlEncode(rsGetDetails.Fields.Item("Length").Value)%>
    <% end if %>
    &nbsp;&nbsp;
    <% if rsGetDetails.fields.item("ProductDetail1").value <> "" then%>
        <%= Server.HTMLEncode(rsGetDetails.Fields.Item("ProductDetail1").Value)%>
    <% end if %>

</td>
</tr>
    <% rsGetDetails.movenext()
        wend 
        %>
    </tbody>
    </table>
    <%= details_message %>
    <%         rsGetProducts.MoveNext()
If rsGetProducts.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
        %>
<!--#include file="includes/inc-pull-products-paging.asp" -->
<% end if ' rsGetProducts.eof%>
</div>
</body>
<script type="text/javascript" src="../js/jquery-3.3.1.min.js"></script>
<script>
	// Confirm item pulled
	$(".confirm-item-pulled").click(function () {
        detailid = $(this).attr("data-detailid");
        productid = $(this).attr("data-productid");
        qty_counted = $('#qty_' + detailid).val();

        if ($('.check-' + detailid).hasClass("text-black-50")) {
            status = "update"
        } else {
            status = "clear"
        }
        console.log(status);

        $.ajax({
        method: "post",
        dataType: "json",
		url: "ajax/ajax-pull-item-qty-count.asp",
		data: {detailid: detailid, productid: productid, qty_counted: qty_counted, status: status}
		})
		.done(function(json,msg) {
            $('.check-' + detailid).toggleClass("text-black-50 text-success");
            $('.tr-' + detailid).toggleClass("table-success");
            $('#qty_' + detailid).prop('disabled', function(i, v) { return !v; });

            if (json.status == "complete") {
                $('#frm-scan-to-bin').show();
                $('html,body').scrollTop(0);
            } else {
                $('#frm-scan-to-bin').hide();
            }
		})
		.fail(function(json,msg) {
			alert("ERROR");
        });
    });
    
    
            // Assign product to user
	$("#btn-assign").click(function () {
        productid = $(this).attr("data-productid");

        $.ajax({
        method: "post",
		url: "ajax/ajax-assign-to-packer.asp",
		data: {productid: productid}
		})
		.done(function(msg) {
            location.reload();
		})
		.fail(function(msg) {
			alert("ERROR");
        });  
    });

    // Finalize product into limited bin
	$("#btn-finalize").click(function () {
        bin_number = $('#bin_number').val();
        productid = $(this).attr("data-productid");

        $.ajax({
        method: "post",
		url: "includes/inc-assign-discontinued.asp",
		data: {bin: bin_number, productid: productid}
		})
		.done(function(msg) {
            $('#message').html('<div class="alert alert-success">SUCCESS - Products assigned</div>').show();
            $('tbody, #frm-scan-to-bin').hide();
		})
		.fail(function(msg) {
			$('#message').html('<div class="alert alert-danger">Error processing</div>').show();
        });  
    });


        // Reset / remove item to be pulled
	$("#btn-delete").click(function () {
        bin_number = $('#bin_number').val();
        productid = $(this).attr("data-productid");

        $.ajax({
        method: "post",
		url: "ajax/ajax-remove-item-topull.asp",
		data: {productid: productid}
		})
		.done(function(msg) {
            $('#message').html('<div class="alert alert-success">SUCCESS</div>').show();
            location.reload();
		})
		.fail(function(msg) {
			$('#message').html('<div class="alert alert-danger">Error</div>').show();
        });  
    });

</script>
</html>
