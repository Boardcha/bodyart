<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT po_qty, po_detailid FROM tbl_po_details WHERE po_orderid = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10, request.form("po_id") ))
set rsGetItems = objCmd.Execute()

While NOT rsGetItems.EOF 

        '===== DEDUCT QUANTITY FROM INVENTORY
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - ? WHERE ProductDetailID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,15, rsGetItems("po_qty") ))
        objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, rsGetItems("po_detailid") ))
        objCmd.Execute()

        '====== WRITE TO EDIT LOGS =======================
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (?, " & rsGetItems("po_detailid") & ",'Automated - Deducted " & rsGetItems("po_qty") & " via manager finalizing bulk order','" & now() & "')"
        objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
        objCmd.Execute()
        Set objCmd = Nothing

rsGetItems.MoveNext()
Wend

'===== TAKE OFF FLAG FOR NEED MANAGER REVIEWED NOW THAT ITS BEEN APPROVED 
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE TBL_PurchaseOrders SET po_needs_review = 0 WHERE PurchaseOrderID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,15, request.form("po_id") ))
        objCmd.Execute()

Set rsGetUser = nothing
DataConn.Close()
%>