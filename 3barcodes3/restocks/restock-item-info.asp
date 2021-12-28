
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("item") <> "" then

    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT d.ProductID, d.ProductDetailID, ISNULL(d.Gauge, '') + ' ' + ISNULL(d.Length, '') + ' ' + ISNULL(d.ProductDetail1, '') + ' ' + j.title AS 'description', j.largepic, d.qty, CASE WHEN d .BinNumber_Detail <> 0 THEN CASE WHEN d.qty > 0  AND new_page_date < (getdate() - 14) THEN 'BIN # ' + CAST(BinNumber_Detail AS varchar(10)) + '  - ITEMS ALREADY IN BIN - BAG # ' + ISNULL(CAST(d .location AS varchar(10)), '') ELSE 'BIN # ' + CAST(BinNumber_Detail AS varchar(10)) + ' - NEW BAG IN BIN' END ELSE ISNULL(loc.ID_Description, '') + ' ' + ISNULL(CAST(d .location AS varchar(10)), '') END AS 'location' FROM jewelry AS j INNER JOIN ProductDetails AS d ON j.ProductID = d.ProductID INNER JOIN TBL_Barcodes_SortOrder AS loc ON d.DetailCode = loc.ID_Number WHERE d.ProductDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
    set rsGetItem = objCmd.Execute()

    ' ====== GET INFORMATION FOR QTY TO STORE IN EDITS LOG
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT qty FROM ProductDetails WHERE ProductDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
    set rsGetCurrentStock = objCmd.Execute()

    ' ====== GET INFORMATION FOR PURCHASE ORDER THE ITEM IS IN
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT po_qty FROM tbl_po_details WHERE po_detailid = ? AND po_orderid = ? AND po_received = 0"
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
    objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,20, request.form("po_id")))
    set rsPurchaseOrder = objCmd.Execute()

    '======= IF AN ACTIVE ITEM IS FOUND IN PURCHASE ORDER THEN ADD QTY BACK IN STOCK, OTHERWISE DO NOT THING -- AVOIDS DOUBLE SCANNING OR PUTTING ITEMS BACK IN STOCK MORE THAN ONCE ======================
    if request.form("po_id") <> 0 then
        if NOT rsPurchaseOrder.eof then

            Set objCmd = Server.CreateObject ("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "UPDATE tbl_po_details SET po_received = 1, po_date_received = '" & now() & "' WHERE po_detailid = ? AND po_orderid = ? AND po_received = 0"
            objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
            objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,20, request.form("po_id")))
            objCmd.Execute 

            Set objCmd = Server.CreateObject ("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "UPDATE ProductDetails SET qty = qty + ? WHERE ProductDetailID = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,20, rsPurchaseOrder.Fields.Item("po_qty").Value ))
            objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
            objCmd.Execute 

            var_qty_log_text = "Scanned restock - Updated qty from " & Cint(rsGetCurrentStock.Fields.Item("qty").Value) & " to " & Cint(rsPurchaseOrder.Fields.Item("po_qty").Value) + Cint(rsGetCurrentStock.Fields.Item("qty").Value)

        else 
            var_qty_log_text = "Scanned restock -  Duplicate item scan, did not add qtys again"
        end if '=== IF A PURCHASE ORDER IS FOUND
    else '=== IF THE PURCHASE ORDER SCAN IS A 0 -- IF AN ITEM DOESN'T HAVE A PURCHASE ORDER, IT STILL NEEDS TO BE PUT IN STOCK. THESE ARE THINGS LIKE CUSTOMER RETURNS OR LOST AND FOUND ITEMS.=====================

            var_qty_log_text = "Scanned restock -  Qty's were manually adjusted. Automated system did not adjust qty."
        
    end if '=== IF THE PURCHASE ORDER IS NOT 0

    '====== UPDATE EDITS LOG WITH QTY UPDATE INFORMATION ===================
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, edit_date, product_id, detail_id, description) VALUES(?, GETDATE(), ?, ?, ?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("product_id",3,1,15, rsGetItem.Fields.Item("ProductID").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("detail_id",3,1,15, request.form("item")))
    objCmd.Parameters.Append(objCmd.CreateParameter("var_qty_log_text",200,1,100, var_qty_log_text))
    objCmd.Execute 

end if '==== if info in scanned field is present

%>

<!DOCTYPE html>
<html lang="en">
<body>

<% if NOT rsGetItem.eof then %>
<div class="alert alert-info h5"><%=rsGetItem.Fields.Item("location").Value  %></div>
<div class="h6 mt-2"><%=rsGetItem.Fields.Item("description").Value  %></div>
            <img class="img-fluid" src="http://bodyartforms-products.bodyartforms.com/<%= rsGetItem.Fields.Item("largepic").Value %>">
<% else %>
    <div class="alert alert-danger mt-3">No item found</div>
<% end if %>
</body>
</html>