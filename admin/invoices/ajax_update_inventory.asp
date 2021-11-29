<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT qty, DetailID, InvoiceID, title, ProductDetail1 FROM QRY_OrderDetails  WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,request("invoiceid")))
	set rsUpdate = objCmd.Execute()

While NOT rsUpdate.EOF

	if request("type") = "add" then
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET qty = qty + " & rsUpdate.Fields.Item("qty").Value & ", active = 1 WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
		objCmd.Execute()
		
		' Reactive main product listing
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn  
		objCmd.CommandText = "UPDATE jewelry SET active = 1 FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
		objCmd.Execute()
		
		'Write info to edits log	
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (?, " & rsUpdate("DetailID") & ",'Automated - Added " & rsUpdate("qty") & " to qty using add quantities button on invoice page','" & now() & "')"
		objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
		objCmd.Execute()
		Set objCmd = Nothing
		
	else
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & rsUpdate.Fields.Item("qty").Value & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
		objCmd.Execute()

			
		'Write info to edits log	
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (?, " & rsUpdate("DetailID") & ",'Automated - Deducted " & rsUpdate("qty") & " from qty using deduct quantities button on invoice page','" & now() & "')"
		objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
		objCmd.Execute()
		Set objCmd = Nothing

	end if
	

	rsUpdate.MoveNext()
Wend

	if request("type") = "add"  then
		var_type = "Items have been put back into stock via the add quantities button"
	else
		var_type = "Items have been deducted from stock via the deduct quantities button"
	end if

	'===== INSERT NOTE ABOUT AUTOMATED UPDATE ================
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,request("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,"Automated message - " & var_type ))
	objCmd.Execute()

DataConn.Close()
%>