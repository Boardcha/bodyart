
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("item") <> "" then

    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT d.ProductID, d.ProductDetailID, ISNULL(d.Gauge, '') + ' ' + ISNULL(d.Length, '') + ' ' + ISNULL(d.ProductDetail1, '') + ' ' + j.title AS 'description', j.largepic, d.qty, CASE WHEN d .BinNumber_Detail <> 0 THEN CASE WHEN d.qty > 0 THEN 'BIN # ' + CAST(BinNumber_Detail AS varchar(10)) + ISNULL(CAST(d .location AS varchar(10)), '') ELSE 'BIN # ' + CAST(BinNumber_Detail AS varchar(10)) END ELSE ISNULL(loc.ID_Description, '') + ' ' + ISNULL(CAST(d .location AS varchar(10)), '') END AS 'location' FROM jewelry AS j INNER JOIN ProductDetails AS d ON j.ProductID = d.ProductID INNER JOIN TBL_Barcodes_SortOrder AS loc ON d.DetailCode = loc.ID_Number WHERE d.ProductDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
    set rsGetItem = objCmd.Execute()

    ' ====== INSERT EDITS LOG WITH SCAN INFORMATION
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, edit_date, product_id, detail_id, description, po_detailid) VALUES(?, GETDATE(), ?, ?, 'Tagged product for restocking', ?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("product_id",3,1,15, rsGetItem.Fields.Item("ProductID").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("detail_id",3,1,15, request.form("item")))
    objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,20, request.form("po_id")))
    objCmd.Execute 

end if '==== if info in scanned field is present

%>

<!DOCTYPE html>
<html lang="en">
<body>

<% if NOT rsGetItem.eof then %>
<div class="alert alert-info h6">
    <%=rsGetItem.Fields.Item("location").Value  %><br/>
    <%=rsGetItem.Fields.Item("description").Value  %></div>
            <img class="img-fluid" src="http://bodyartforms-products.bodyartforms.com/<%= rsGetItem.Fields.Item("largepic").Value %>">
<% else %>
    <div class="alert alert-danger mt-3">No item found</div>
<% end if %>
</body>
</html>