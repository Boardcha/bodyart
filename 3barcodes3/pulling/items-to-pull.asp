
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("invoices") <> "" then
var_orderissue = 0

' ====== GET ORDER
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ID,  customer_comments, customer_first, customer_last, sent_items.item_description, sent_items.shipped, sent_items.autoclave, sent_items.PackagedBy,  pulled_by FROM sent_items WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,20, request.form("invoices")))
Set rsGetOrder = objCmd.Execute()

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TBL_OrderSummary.InvoiceID, TBL_OrderSummary.OrderDetailID, jewelry.ProductID, ProductDetails.ProductDetailID, ProductDetails.location,  TBL_OrderSummary.qty, TBL_OrderSummary.item_price, jewelry.title, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1, jewelry.picture, img_thumb, jewelry.customorder, jewelry.pair, ProductDetails.qty AS AmtInStock, jewelry.type, anodization_fee,  ProductDetails.BinNumber_Detail, ProductDetails.Free_QTY, TBL_Barcodes_SortOrder.ID_Number, TBL_Barcodes_SortOrder.ID_Description, TBL_Barcodes_SortOrder.ID_SortOrder, TBL_OrderSummary.TimesScanned, TBL_OrderSummary.ScanItem_Timestamp FROM TBL_OrderSummary INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number LEFT OUTER JOIN tbl_images ON  ProductDetails.img_id = tbl_images.img_id WHERE (jewelry.ProductID <> 3704) AND (jewelry.ProductID <> 3611) AND (jewelry.ProductID <> 3587) AND (jewelry.ProductID <> 2890)  AND (jewelry.ProductID <> 3704) AND (jewelry.ProductID <> 3612) AND (jewelry.ProductID <> 3143) AND (jewelry.ProductID <> 3144) AND (jewelry.ProductID <> 10171) AND (jewelry.ProductID <> 6733) AND (jewelry.ProductID <> 3146) AND (jewelry.ProductID <> 3145) AND InvoiceID = ? ORDER BY TBL_OrderSummary.TimesScanned DESC, CASE WHEN customorder = 'yes' THEN 1 ELSE 2 END ASC, anodization_fee ASC, ProductDetails.BinNumber_Detail ASC, TBL_Barcodes_SortOrder.ID_SortOrder, ID_Description ASC, location ASC, ProductDetails.ProductDetailID"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,20, request.form("invoices")))

set rsGetItems = Server.CreateObject("ADODB.Recordset")
rsGetItems.CursorLocation = 3 'adUseClient
rsGetItems.Open objCmd

' ====== UPDATE ORDER WITH PULLER INFO, DATE STARTED
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE sent_items SET pulled_by = ?, ScanInvoice_Timestamp = GETDATE() WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("pulled_by",200,1,50, rsGetUser.Fields.Item("name").Value ))
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15, request.form("invoices") ))
objCmd.Execute  

' Check if autoclave service is in the products 
if NOT rsGetItems.eof then
    while not rsGetItems.eof
        if rsGetItems.Fields.Item("ProductID").Value = 1464 then
            autoclave = "yes"
        end if
    rsGetItems.movenext()
    wend
    rsGetItems.moveFirst()
end if '==== rsGetItems.eof
%>

<!DOCTYPE html>
<html lang="en">
<style>
.table tbody tr:nth-child(4n+1), .table tbody tr:nth-child(4n+2) {
 background: #F2F2F2;
}

.table tbody tr:nth-child(8n+1), .table tbody tr:nth-child(8n+2) {
 background: #F2F2F2;
}
</style>
<body>
        <%
        if InStr(rsGetOrder.Fields.Item("item_description").Value, "ORDER UPDATED") > 0 then
        %>
        <div class="alert alert-danger m-2">
                <h5 class="alert-heading my-1">ORDER HAS BEEN UPDATED</h5>
                Invoice needs to be reprinted</div>
        <%
        end if

        if InStr(rsGetOrder.Fields.Item("item_description").Value, "ADDRESS UPDATED") > 0 then
        %>
        <div class="alert alert-danger m-2">
                <h5 class="alert-heading my-1">ADDRESS UPDATED</h5>
                Reprint shipping label</div>
        <%
        end if

        if InStr(rsGetOrder.Fields.Item("item_description").Value, "SHIPPING METHOD UPDATED") > 0 then
        %>
        <div class="alert alert-danger m-2">
                <h5 class="alert-heading my-1">SHIPPING METHOD UPDATED</h5>
                Reprint BOTH shipping label &amp; invoice</div>
        <%
        end if

        if rsGetOrder.Fields.Item("shipped").Value = "ON HOLD" then
        %>
        <div class="alert alert-danger m-2">
                <h5 class="alert-heading my-1">ORDER ON HOLD</h5>
                Set aside and check later</div>
        <%
        end if

        if rsGetOrder.Fields.Item("shipped").Value = "Cancelled" then
        %>
        <div class="alert alert-danger m-2">
                <h5 class="alert-heading my-1">ORDER CANCELLED</h5>
                Refund the shipping label and shred invoice
                </div>
        <%
        end if
        %>
            <% if instr(lcase(rsGetOrder.Fields.Item("item_description").Value),lcase("CONSERVE")) > 0 then %>
            <i class="fa fa-recycle fa-2x text-success pl-2 pr-1 pt-1 d-inline-block align-top"></i>
            <% end if %>
    <% if autoclave = "yes" then %>
        <i class="fa fa-first-aid fa-2x text-primary px-1 d-inline-block align-top"></i>
    <% END IF %>
    <%
    if rsGetOrder("item_description") <> "" then 
        var_public_comments = REPLACE(rsGetOrder("item_description"), "<b><font size=3>BACKORDER SHIPMENT</font></B><br>", "")
        var_public_comments = REPLACE(var_public_comments, "CONSERVE PLASTIC BAGS<br>", "")
        var_public_comments = REPLACE(var_public_comments, "ORDER UPDATED", "")
        var_public_comments = REPLACE(var_public_comments, "ADDRESS UPDATED", "")
        var_public_comments = REPLACE(var_public_comments, "SHIPPING METHOD UPDATED", "")
        var_public_comments = REPLACE(var_public_comments, "GIFT ORDER<br>", "")
    
    %>
    <div class="d-inline-block small text-primary p-0 m-0 align-top">
        <%= var_public_comments %>
    </div>
    <% end if %>
<% if NOT rsGetItems.eof then %>
<table class="table tabl-sm">
    <% if var_orderissue = 0 then
    while not rsGetItems.eof 

    if rsGetItems.Fields.Item("ProductID").Value <> 1464 then 'not autoclave
    
    ' For items that always need lots of scanning, bypass and just allow the completed info to display
    If rsGetItems.Fields.Item("ProductID").Value = 530 OR rsGetItems.Fields.Item("ProductID").Value = 1649 OR rsGetItems.Fields.Item("ProductID").Value = 15385 then
        var_matchqty = "yes"
    else
        var_matchqty = "no"
    end if

    If rsGetItems.Fields.Item("ProductID").Value = 2180 OR rsGetItems.Fields.Item("ProductID").Value = 25324 OR rsGetItems.Fields.Item("BinNumber_Detail").Value <> 0 OR rsGetItems.Fields.Item("customorder").Value = "yes" then
        var_location = rsGetItems.Fields.Item("ProductDetailID").Value
    elseif  rsGetItems("anodization_fee") > 0 then
        var_location = rsGetItems("OrderDetailID")
    else
        var_location = rsGetItems.Fields.Item("ID_Number").Value & rsGetItems.Fields.Item("location").Value
    end if

    if rsGetItems.Fields.Item("img_thumb").Value <> "" then
        var_thumbnail = rsGetItems.Fields.Item("img_thumb").Value
    else
        var_thumbnail = rsGetItems.Fields.Item("picture").Value
    end if
    %>
	<%
	TimesScanned = rsGetItems("TimesScanned")
	If TimesScanned >= rsGetItems("qty") Then 
		TimesScanned = rsGetItems("qty")
		ItemScanCompleted = true
	Else
		ItemScanCompleted = false
	End If
	%>	
    <% if var_previous_detailid = rsGetItems.Fields.Item("ProductDetailID").Value AND rsGetItems("anodization_fee") = 0 then 
        var_addqty = rsGetItems.Fields.Item("qty").Value
    %>
<script type="text/javascript">
    var duplicate_detailid = $("tr[data-productdetailid='" + <%= rsGetItems.Fields.Item("productdetailid").Value %> + "']:first").attr('data-orderdetailid');
    var duplicate_qty = $("tr[data-productdetailid='" + <%= rsGetItems.Fields.Item("productdetailid").Value %> + "']:first").attr('data-qty');
    var new_qty = parseInt(duplicate_qty) + parseInt(<%= var_addqty %>)

    $('#duplicate_' + duplicate_detailid).html(new_qty);

    $("tr[data-productdetailid='" + <%= rsGetItems.Fields.Item("productdetailid").Value %> + "']:first").attr('data-qty', new_qty);
</script>
    <% else  '===var_previous_detailid %>
    <tr class="item-information <%If TimesScanned > 0 AND TimesScanned < rsGetItems("qty") Then Response.Write "table-warning"%>" id="<%= rsGetItems.Fields.Item("OrderDetailID").Value %>" data-invoice="<%= rsGetItems.Fields.Item("InvoiceID").Value %>" data-productdetailid="<%= rsGetItems.Fields.Item("ProductDetailID").Value %>" data-orderdetailid="<%= rsGetItems.Fields.Item("OrderDetailID").Value %> "data-location="<%= var_location %>" data-qty="<%= rsGetItems.Fields.Item("Qty").Value %>" data-matchqty="<%= var_matchqty %>" data-timescanned="0" data-status="not fullfilled">
        <td class="py-0 align-middle pl-0 pr-1" style="width:50px">
            <img src="http://bodyartforms-products.bodyartforms.com/<%= var_thumbnail %>" class="expand" data-orderdetailid="<%= rsGetItems.Fields.Item("OrderDetailID").Value %>" style="width:50px;height:auto">
        </td>
        <td class="py-0 align-middle">
            <% 
                'response.write "<br>DEBUG - # SCAN " & var_location
            if rsGetItems("anodization_fee") = 0   then 
                'response.write "<br>DEBUG - No anodization needed" 
            end if
                %>
                <span class="alert alert-secondary py-0 px-1 font-weight-bold">
                  <%
                    if rsGetItems.Fields.Item("Free_QTY").Value > 0 AND rsGetItems.Fields.Item("item_price").Value <= 0 then
        %>
                FREE
        <%
        end if
        
         if rsGetItems.Fields.Item("ID_Description").Value <> "Main" AND rsGetItems.Fields.Item("ID_Description").Value <> "Free" AND rsGetItems.Fields.Item("BinNumber_Detail").Value = 0 AND rsGetItems("anodization_fee") = 0  then %>
                <span class="mr-1 border-secondary border-right">
                    <%= rsGetItems.Fields.Item("ID_Description").Value %>
                </span>
            <% end if %>
        
        <%
        BinType = ""		 
        If rsGetItems.Fields.Item("BinNumber_Detail").Value <> 0 then
             
            If  (rsGetItems.Fields.Item("ID_Description").Value = "Case 1" OR rsGetItems.Fields.Item("ID_Description").Value = "Case 2" OR rsGetItems.Fields.Item("ID_Description").Value = "Case 3" OR rsGetItems.Fields.Item("ID_Description").Value = "Case 4") Then  
                BinType = "<span class='mr-1 pr-1 border-secondary border-right'>" & rsGetItems("ID_Description") & "</span>Shelf"
            else
                BinType = "BIN"
            end if
        %>
                
                <%= BinType %>
                <span class="ml-1 mr-1 border-secondary border-right">
                    <%= rsGetItems.Fields.Item("BinNumber_Detail").Value %> 
                </span>
           
        <%
        End if '--- if bin # is not zero
    
                
        if rsGetItems.Fields.Item("customorder").Value <> "yes" AND rsGetItems("anodization_fee") = 0 then 
            If rsGetItems.Fields.Item("BinNumber_Detail").Value <> 0 then
            '==== Show detail id for items in limited bins
            %>   
                <%= rsGetItems.Fields.Item("ProductDetailID").Value %>
            <% else 
            '===== regular stock item location
            %>
            <%= rsGetItems.Fields.Item("location").Value %>
            
            <% end if
            
            ' ===== If pick random colored sticker, show customer name 
            if rsGetItems.Fields.Item("ProductDetailID").Value = "72198" then %>   
                 (<%= rsGetOrder.Fields.Item("customer_first").Value %>)
            <% end if %>
        </span>
        <% if rsGetItems.Fields.Item("BinNumber_Detail").Value <> 0 AND rsGetItems("ProductDetailID") <> rsGetItems("location") then
        %> 
        <span class="d-block ml-1">Old location <%= rsGetItems("location") %></span>
        <% end if '  also show old location for limited bin labels IF DETAIL ID DOESNT MATCH %>
        <% else %>
            <% if rsGetItems("customorder") = "yes" then %>
            CUSTOM ORDER FOR
            <% end if %>
            <% if rsGetItems("anodization_fee") > 0 then %>
            CUSTOM ANODIZATION FOR
            <% end if %>
             | <%= rsGetOrder.Fields.Item("customer_first").Value %></span>
            <div>Invoice # <%= rsGetOrder.Fields.Item("ID").Value %></div>
        <% end if %> 
        <% ' If a returned package is in the order
        if rsGetItems.Fields.Item("ProductID").Value = 25087 then
        %>
        <div class="alert alert-info">
        Ship returned order on shelf for 
        <div class="font-weight-bold"><%= rsGetOrder.Fields.Item("customer_first").Value %>&nbsp;<%= rsGetOrder.Fields.Item("customer_last").Value %> -- Invoice # <%= rsGetOrder.Fields.Item("ID").Value %></div>
        </div>
        <%
        end if
        %>
        <% ' If a returned package is in the order
        if rsGetItems.Fields.Item("ProductID").Value = 2991 then
        %>
        <div class="alert alert-info">
        Send a return mailer
        </div>
        <%
        end if
        %>
        </td>
        <td class="py-0 align-middle">
            <span class="alert alert-success py-0 px-1 h5 font-weight-bold">
                    <span class="mr-1 pr-1 border-success border-right">QTY</span><span id="still_need_<%= rsGetItems.Fields.Item("OrderDetailID").Value %>"><%=TimesScanned%></span> of <span class="mr-1" id="duplicate_<%= rsGetItems.Fields.Item("OrderDetailID").Value %>"><%= rsGetItems.Fields.Item("Qty").Value %></span><% if  rsGetItems.Fields.Item("pair").Value = "yes" then %>pair<% end if %>
            </span>
            <i class="fa fa-check-circle fa-lg ml-1 scan_complete toggle-done <%If ItemScanCompleted = true Then Response.Write "text-success" %>" style="color:#BDBDBD" id="check_<%= rsGetItems.Fields.Item("OrderDetailID").Value %>" data-orderdetailid="<%= rsGetItems.Fields.Item("OrderDetailID").Value %>"></i>
        </td>
    </tr>
    <tr id="<%= rsGetItems.Fields.Item("OrderDetailID").Value %>_sub" <%If TimesScanned = 0 Then %>style="display:none"<%End If%>>    
        <td class="border-top-0 pt-0 align-top <%If TimesScanned > 0 AND TimesScanned < rsGetItems("qty") Then Response.Write "table-warning"%>" style="border-bottom: 3px solid grey" colspan="3">
                <%= rsGetItems.Fields.Item("title").Value %>&nbsp;<%= rsGetItems.Fields.Item("ProductDetail1").Value %>&nbsp;<%= rsGetItems.Fields.Item("Gauge").Value %>&nbsp;<%= rsGetItems.Fields.Item("Length").Value %><br>
                <span class="alert alert-info py-0 px-1 font-weight-bold"><span class="mr-1 pr-1 border-info border-right bo-button"  data-toggle="modal" data-target="#modal-submit-backorder" data-orderdetailid="<%= rsGetItems.Fields.Item("OrderDetailID").Value %>">BO</span>In stock: <%= rsGetItems.Fields.Item("AmtInStock").Value %></span>

                <span class="alert alert-warning py-0 px-1 font-weight-bold text-dark ml-3 error-button" style="color:#BDBDBD"   data-toggle="modal" data-target="#modal-submit-error"  data-orderdetailid="<%= rsGetItems.Fields.Item("OrderDetailID").Value %>">Report issue</span>
        </td>
    </tr>
    <% end if ' var_previous_detailid %>
    <%
    end if ' not autoclave

    var_addqty = 0

    '==== ONLY COMBINE ITEMS IF THEY ARE NOT HAVING CUSTOM COLOR SERVICE ADDED ON
    if rsGetItems("anodization_fee") = 0   then 
        var_previous_detailid = rsGetItems.Fields.Item("ProductDetailID").Value
    else
        var_previous_detailid = rsGetItems("OrderDetailID")
    end if

    
    rsGetItems.movenext()
    wend
    end if ' === var_orderissue = 0
    %>
</table>
<% end if '==== rsGetItems.eof %>
</body>
</html>
<% else %>
    <div class="alert alert-danger mt-3">No invoices selected</div>
<% end if %>
