<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	column_title = request.form("column")
	
	' protection to stop page from sql injection
	if Len(column_title) > 25 then
		response.end
	end if

	
	column_value = request.form("value")
	id = request.form("id")
	friendly_name = request.form("friendly_name")

if request.form("detailid") = "" then ' check to see if a detailid is provided and if not, just update the main products table	

	productdetailid = 0
	productid = 0

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "select " & column_title & " from sent_items where ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	set rsGetColumn = objCmd.Execute()
	
	orig_value = rsGetColumn.Fields.Item(column_title).Value
	edits_description = "Updated MAIN INVOICE: " & friendly_name & " from " & orig_value & " to " & column_value
	edits_detail_id = 0

	response.write column_title
	response.write column_value
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update sent_items set " & column_title & " = ? where ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,8000,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	
else ' update the details

	detailid = request.form("detailid")
	productdetailid = request.form("productdetailid")
	productid = request.form("productid")

	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "select " & column_title & ", OrderDetailID from TBL_OrderSummary where OrderDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,detailid))
	set rsGetColumn = objCmd.Execute()
	
	orig_value = rsGetColumn.Fields.Item(column_title).Value
	edits_description = "Updated INVOICE ITEM: " & friendly_name & " from " & orig_value & " to " & column_value
	edits_detail_id = rsGetColumn.Fields.Item("OrderDetailID").Value

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update TBL_OrderSummary set " & column_title & " = ? where OrderDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,225,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,detailid))
	objCmd.Execute()

end if


		'Write ALL info to edits log table
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, invoice_id, invoice_detail_id, detail_id, product_id, description, edit_date) VALUES (" & user_id & "," & id & "," & edits_detail_id & "," & productdetailid & "," & productid & ",'" & edits_description & "','" & now() & "')"
		objCmd.Execute()
	

DataConn.Close()
%>