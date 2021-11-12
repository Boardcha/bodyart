<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' Create new empty order if move to invoice # is 0
if lCase(request("move_to_id")) = "new" then

		Set objCmd = Server.CreateObject ("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT ID, email, country FROM sent_items WHERE ID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,request("invoiceid")))
		Set rsGetOldInvoice = objCmd.Execute()

		if rsGetOldInvoice.Fields.Item("country").Value <> "USA"  AND GetNotes("country") <> "US"  then
			shipping = "DHL GlobalMail Packet Priority"
		else
			shipping = "DHL Basic mail"
		end if
		 
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO sent_items (shipped, customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, date_order_placed, shipping_rate, shipping_type, ship_code, phone, pay_method, UPS_Service, autoclave) SELECT 'Pending...', customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, '" & now() & "',0 ,'" & shipping & "' , 'paid', phone, pay_method, '', autoclave FROM sent_items WHERE ID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,request("invoiceid")))
		objCmd.Execute() 
		
		Set objCmd = Server.CreateObject ("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TOP(1) ID FROM sent_items WHERE email = ? ORDER BY ID DESC" 
		objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,75,rsGetOldInvoice.Fields.Item("email").Value))
		Set rsGetNewestInvoice = objCmd.Execute()		

	move_to_invoice = rsGetNewestInvoice.Fields.Item("ID").Value
else
	move_to_invoice = request("move_to_id")
end if


'move detail row to a new invoice
if request.form("toggle_type") = "move" then

		detail_array =split(request.form("details"),",")
		For Each strItem In detail_array

			if strItem <> "" then 
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE TBL_OrderSummary SET InvoiceID = ? WHERE OrderDetailID = " + strItem
				objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,move_to_invoice))
				objCmd.Execute()
			end if ' make sure a detail id is provided to write db
			
		Next
%>
{  
   "invoiceid":"<%= move_to_invoice %>"
}
<%	
end if ' move detail to a new invoice

'copy detail row
if request.form("toggle_type") = "copy" then

'response.write request.form("details")
		' break out form variables into details and rebuild WHERE statement
		detail_array =split(request.form("details"),",")
		For Each strItem In detail_array
	
		'response.write strItem + " ID "

			if strItem <> "" then 
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "INSERT INTO TBL_OrderSummary(InvoiceID, ProductID, DetailID, qty, item_price) SELECT ? , ProductID, DetailID, qty, item_price FROM TBL_OrderSummary WHERE OrderDetailID = " + strItem 
				objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,move_to_invoice))
				objCmd.Execute()

				' Get information to deduct inventory
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "SELECT qty, DetailID FROM QRY_OrderDetails WHERE OrderDetailID = " + strItem 
				objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,request("invoiceid")))
				set rsUpdate = objCmd.Execute()
				
				' Deduct inventory
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & rsUpdate.Fields.Item("qty").Value & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
				objCmd.Execute()
							
			end if ' make sure a detail id is provided to write db
			
		Next

%>
{  
   "invoiceid":"<%= move_to_invoice %>"
}
<%		
end if ' if copying detail into a new product


' If a return mailer was checked to be sent
	if request("return_mailer") = "yes" then 
	
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_OrderSummary(InvoiceID, ProductID, DetailID, qty, item_price) VALUES (?, 2991, 17999, 1, 0)"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,move_to_invoice))
		objCmd.Execute()
							
	end if ' make sure a detail id is provided to write db
	
' If a reship package was checked to be sent
	if request("reship_returned") = "yes" then 
	
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_OrderSummary(InvoiceID, ProductID, DetailID, qty, item_price) VALUES (?, 25087, 143192, 1, 0)"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,move_to_invoice))
		objCmd.Execute()
							
	end if ' make sure a detail id is provided to write db

DataConn.Close()
%>