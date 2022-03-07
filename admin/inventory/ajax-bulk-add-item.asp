<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
    '==== ADD ROW TO PURCHASE ORDER TABLE
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_po_details (po_temp_id, po_detailid, po_qty) VALUES (?, ? , ?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("var_temp_po_id",3,1,20, request.Cookies("bulk-po-id") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, request("detailid") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,15, request("qty") ))
    objCmd.Execute()

    '====== ONLY DEDUCT INVENTORY IF A MANAGER DOES NOT NEED TO REVIEW THE FINAL ORDER TO BE PULLED =====
    if request.form("var_needs_review") = "no" then
        '===== DEDUCT QUANTITY FROM INVENTORY SO IT CAN BE PULLED BY STAFF
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - ? WHERE ProductDetailID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,15, request("qty") ))
        objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, request("detailid") ))
        objCmd.Execute()

        '====== WRITE TO EDIT LOGS =======================
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (?, " & request("detailid") & ",'Automated - Deducted " & request("qty") & " adding item to bulk purchase order','" & now() & "')"
        objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
        objCmd.Execute()
        Set objCmd = Nothing

    end if

Set rsGetUser = nothing
DataConn.Close()
%>