<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO TBLWaitingList (DetailID, email, customerID, date_added, waiting_qty) VALUES (?, ?, ?, ?, ?)"
objCmd.Parameters.Append(objCmd.CreateParameter("DetailID",3,1,10, request.form("detailid")))
objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,75, request.form("waiting-email")))
objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
objCmd.Parameters.Append(objCmd.CreateParameter("DateAdded",200,1,12,date()))
objCmd.Parameters.Append(objCmd.CreateParameter("waiting-qty",3,1,10, request.form("waiting-qty")))
objCmd.Execute()


if request.form("wishlist_id") <> "" then

	' notate wishlist item that it's been added to waiting list
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE wishlist SET waiting_list = 1 WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("wishlist_id",3,1,12,request.form("wishlist_id")))
	objCmd.Execute()


end if


DataConn.Close()
Set DataConn = Nothing
%>