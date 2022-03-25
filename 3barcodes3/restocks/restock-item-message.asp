
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("bin") <> "" then

    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT d.ProductID, d.ProductDetailID, ISNULL(d.Gauge, '') + ' ' + ISNULL(d.Length, '') + ' ' + ISNULL(d.ProductDetail1, '') + ' ' + j.title AS 'description', j.largepic, d.qty, CASE WHEN d .BinNumber_Detail <> 0 THEN CASE WHEN d.qty > 0 THEN 'BIN # ' + CAST(BinNumber_Detail AS varchar(10)) + '  - ITEMS ALREADY IN BIN - BAG # ' + ISNULL(CAST(d .location AS varchar(10)), '') ELSE 'BIN # ' + CAST(BinNumber_Detail AS varchar(10)) + ' - NEW BAG IN BIN' END ELSE ISNULL(loc.ID_Description, '') + ' ' + ISNULL(CAST(d .location AS varchar(10)), '') END AS 'location' FROM jewelry AS j INNER JOIN ProductDetails AS d ON j.ProductID = d.ProductID INNER JOIN TBL_Barcodes_SortOrder AS loc ON d.DetailCode = loc.ID_Number WHERE CASE WHEN d .BinNumber_Detail <> 0 THEN CAST(BinNumber_Detail AS varchar(10)) ELSE CAST(DetailCode as varchar(10))  + CAST(d.location as varchar(10)) END = ? AND d.ProductDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("location",3,1,20, request.form("bin")))
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
    set rsScanStatus = objCmd.Execute() 

end if '==== if info in scanned field is present

%>


{
<% if NOT rsScanStatus.eof then %>
    "status":"match"
<%     ' ====== INSERT EDITS LOG WITH ALL INFORMATION OF LOCATION SCANNED NO MATTER IF IT MATCHED OR NOT -- FOR TRACKING
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, edit_date, product_id, detail_id, description, po_detailid) VALUES(?, GETDATE(), ?, ?, 'MATCHED restock scan - SCANNED INTO LOCATION ' + ?, ?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("product_id",3,1,15, rsScanStatus.Fields.Item("ProductID").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("detail_id",3,1,15, request.form("item")))
    objCmd.Parameters.Append(objCmd.CreateParameter("description",200,1,50, rsScanStatus.Fields.Item("location").Value))
    objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,20, request.form("po_id")))
    objCmd.Execute 

else '==== Scan did not match %>
    "status":"no-match"
<% 

    ' ====== GET INFORMATION FOR PRODUCT THAT WAS SCANNED ON A WRONG SCAN TO RECORD TO EDITS LOG
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT d.ProductID, d.ProductDetailID FROM jewelry AS j INNER JOIN ProductDetails AS d ON j.ProductID = d.ProductID INNER JOIN TBL_Barcodes_SortOrder AS loc ON d.DetailCode = loc.ID_Number WHERE d.ProductDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("item")))
    set rsFailedScannedItem = objCmd.Execute()

    ' ====== INSERT EDITS LOG WITH ALL INFORMATION OF LOCATION SCANNED NO MATTER IF IT MATCHED OR NOT -- FOR TRACKING
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, edit_date, product_id, detail_id, description) VALUES(?, GETDATE(), ?, ?, 'NO MATCH restock scan - SCANNED INTO LOCATION ' + ?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("product_id",3,1,15, rsFailedScannedItem.Fields.Item("ProductID").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("detail_id",3,1,15, request.form("item")))
    objCmd.Parameters.Append(objCmd.CreateParameter("description",200,1,50, request.form("bin")))
    objCmd.Execute 

end if %>
}