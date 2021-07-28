<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	' protection to stop page from sql injection
%>
<!--#include file="../../functions/sql_injection_prevention.asp" -->
<%

' Save all detail ID's to array to update database
Dim arrDetails
i = 0
arrDetails = split(replace(request("detailids"), """", ""), ",")

session_tempid = request.form("tempid")

For Each item In arrDetails
if item <> 0 then

	' delete any prior rows with detail id so that we don't have duplicates of the same item in our order
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM tbl_po_details WHERE po_detailid = ? AND po_temp_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,item))
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
	objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_po_details (po_qty, po_temp_id, po_detailid, po_confirmed, po_manual_adjust) VALUES ((SELECT      (ProductDetails.stock_qty - ProductDetails.qty) as po_qty FROM ProductDetails WHERE ProductDetails.ProductDetailID = ?),?,?,1,0)"
	objCmd.Parameters.Append(objCmd.CreateParameter("select_id",3,1,12,item))
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,20,session_tempid))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,12,item))
	objCmd.Execute()
	
	i = i + 1
end if 'if item <> 0
Next


DataConn.Close()
%>