<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	column_title = request.form("column")
	
	' protection to stop page from sql injection
%>
<!--#include file="../../functions/sql_injection_prevention.asp" -->
<%
	
	column_value = request.form("value")
	id = request.form("id")
	friendly_name = request.form("friendly_name")
	session_tempid = request.form("tempid")
	
'	response.write "Friendly: " & friendly_name & " Column" & 

' if we're adding a qty to a PO
if column_title = "po_qty" then


	' delete any prior rows with detail id so that we don't have duplicates of the same item in our order
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM tbl_po_details WHERE po_detailid = ? AND po_temp_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
	objCmd.Execute()
	 
	' only add a qty if it's not 0 
	if column_value <> 0 then
	
	' add new purchase order item
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_po_details (po_qty, po_temp_id, po_detailid, po_confirmed, po_manual_adjust) VALUES (?,?,?,0,1)"
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	
	' if adding by clicking the grey confirm checkmark, make sure to update the database to say it's confirmed and not manually adjust
	if request.form("confirmed") = "yes" then
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE tbl_po_details SET po_confirmed = 1, po_manual_adjust = 0 where po_temp_id = ? and po_detailid = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	
	end if ' request.form("confirmed") = "yes"
	
	end if ' if column_value <> 0 then

else	
	'Update a detail if it's not product level
	if column_title <> "autoclavable" then
	
	' update any other field as usual
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update ProductDetails set " & column_title & " = ? where ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	
	else ' update a column on the product level

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update jewelry set " & column_title & " = ? where ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("value",3,1,10,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	
	end if ' if not a detail level update
	
end if
	

Set rsGetUser = nothing
DataConn.Close()
%>