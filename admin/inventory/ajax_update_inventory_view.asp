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
if column_title = "po_qty" OR column_title ="po_qty_vendor" then

	' Only add a qty if it's bigger than 0 
	If column_value >= 0 then
	
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM tbl_po_details WHERE po_detailid = ? AND po_temp_id = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
		objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
		Set rsPoDetails = objCmd.Execute()
	
		' If detail id exist, update it. Else, Insert a new record.
		If Not rsPoDetails.EOF Then	
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE tbl_po_details SET " & column_title & " = ? where po_temp_id = ? and po_detailid = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50,column_value))
			objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
			objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
			objCmd.Execute()
		Else
			' add new purchase order item
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_po_details (" & column_title & ", po_temp_id, po_detailid, po_confirmed, po_manual_adjust) VALUES (?,?,?,0,1)"
			objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50,column_value))
			objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
			objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
			objCmd.Execute()			
		End If
		
		' if adding by clicking the grey confirm checkmark, make sure to update the database to say it's confirmed and not manually adjust
		if request.form("confirmed") = "yes" then
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE tbl_po_details SET po_confirmed = 1, po_manual_adjust = 0 where po_temp_id = ? and po_detailid = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,session_tempid))
			objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
			objCmd.Execute()
		end if ' request.form("confirmed") = "yes"	
	End If ' if column_value > 0 then
	 
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