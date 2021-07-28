<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../includes/JSON.asp" -->
<!--#include file="../includes/JSON_UTIL.asp" -->
<%
action = Request.Form("action")
page = Request.Form("page")
rangefilter = Request.Form("filter")
	if rangefilter = "never" then
		sql_range = " HAVING (MAX(ISNULL(pd.DateLastPurchased, '01-01-2000')) <= DATEADD(month, -6, getdate()))  "
	else
		sql_range = " HAVING (MAX(pd.DateLastPurchased) <= DATEADD(month, -" & rangefilter & ", getdate())) "
	end if 
	if page = "" then
		page = 1
	end if

pageSize = 100
offset = pageSize * (page - 1)

Response.AddHeader "Content-Type","application/json;charset=utf-8"

if action = "sale" then

	productid = Request.Form("productId")
	sale = Request.Form("amount")
	user = Request.Form("user")

	If sale = 0 then
		OnSale = "N"
	else
		OnSale = "Y"
	end if

	set rsOriginalRecord = Server.CreateObject("ADODB.Recordset")
	rsOriginalRecord.ActiveConnection = MM_bodyartforms_sql_STRING
	rsOriginalRecord.Source = "SELECT SaleDiscount, OnSale, ProductNotes FROM jewelry WHERE ProductID = " & productid
	rsOriginalRecord.CursorLocation = 3 'adUseClient
	rsOriginalRecord.LockType = 1 'Read-only records
	rsOriginalRecord.Open()

	Notes = rsOriginalRecord.Fields.Item("ProductNotes") & " --- Updated Sale status from " & rsOriginalRecord.Fields.Item("OnSale") & "/" & rsOriginalRecord.Fields.Item("SaleDiscount")
	Notes = Notes & " to " & OnSale & "/" & Sale
	Notes = Notes & " by " & user & " on " & Now()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET OnSale = ?, SaleDiscount = ?, ProductNotes = ? WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("sale",200,1,1,  OnSale))
	objCmd.Parameters.Append(objCmd.CreateParameter("discount",6,1,10,  sale))
	objCmd.Parameters.Append(objCmd.CreateParameter("notes",8,1,8000,  Notes))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, productid ))
	on error resume next
	objCmd.Execute()

	set json_response = jsObject()

	json_response("status") = "OK"
	if err <> 0 then
		json_response("status") = "error"
	end if
	json_response("detail") = err

	json_response.Flush
elseif action = "status" then

	productid = Request.Form("productId")
	status = Request.Form("status")
	user = Request.Form("user")

	set rsOriginalRecord = Server.CreateObject("ADODB.Recordset")
	rsOriginalRecord.ActiveConnection = MM_bodyartforms_sql_STRING
	rsOriginalRecord.Source = "SELECT type, ProductNotes FROM jewelry WHERE ProductID = " & productid
	rsOriginalRecord.CursorLocation = 3 'adUseClient
	rsOriginalRecord.LockType = 1 'Read-only records
	rsOriginalRecord.Open()

	Notes = rsOriginalRecord.Fields.Item("ProductNotes") & " --- Updated status from " & rsOriginalRecord.Fields.Item("type")
	Notes = Notes & " to " & status
	Notes = Notes & " by " & user & " on " & Now()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET type = ?, ProductNotes = ? WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("type",8,1,30,  status))
	objCmd.Parameters.Append(objCmd.CreateParameter("notes",8,1,8000,  Notes))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, productid ))
	on error resume next
	objCmd.Execute()

	set json_response = jsObject()

	json_response("status") = "OK"
	if err <> 0 then
		json_response("status") = "error"
	end if
	json_response("detail") = err

	json_response.Flush
elseif action = "count" then



	set rsCheck = Server.CreateObject("ADODB.Recordset")
	rsCheck.ActiveConnection = MM_bodyartforms_sql_STRING
	sqlString = "SELECT j.ProductID, j.title, j.description, j.picture, j.onsale, j.salediscount, CAST(j.ProductNotes AS VARCHAR(MAX)) as ProductNotes, j.type, MAX(pd.DateLastPurchased) as 'LastPurchaseDate',"
	sqlString = sqlString & "MIN(pd.DateLastPurchased) as 'OldestPurchaseDate' FROM jewelry j inner join dbo.ProductDetails pd "
	sqlString = sqlString & "on j.ProductID = pd.ProductID WHERE j.Active = 1 and j.jewelry not like '%save%' and j.title not like '%pre-order%' "
	sqlString = sqlString & "group by j.ProductID, j.title, j.description, j.picture, j.onsale, j.salediscount, CAST(j.ProductNotes AS VARCHAR(MAX)), j.type"
	sqlString = sqlString & sql_range & " ORDER BY MAX(pd.DateLastPurchased)"
	rsCheck.Source = sqlString
	rsCheck.CursorLocation = 3 'adUseClient
	rsCheck.LockType = 1 'Read-only records
	rsCheck.Open()

	set json_response = jsObject()

	json_response("status") = "OK"
	json_response("count") = rsCheck.RecordCount
	json_response.Flush
elseif action = "variants" then

	set rsCheck = Server.CreateObject("ADODB.Connection")
	rsCheck.Open MM_bodyartforms_sql_STRING
	productids = Request.Form("productIds")
	set json_response = jsObject()


	sqlString = "SELECT pd.ProductID, pd.Gauge, pd.Length, pd.ProductDetail1, pd.DateLastPurchased FROM ProductDetails pd inner join jewelry j "
	sqlString = sqlString & "on j.ProductID = pd.ProductID where j.ProductID IN (" 
	sqlString = sqlString & Join(Split(productids, "|"), ",") & ") "

	QueryToJSON(rsCheck, sqlString).Flush
else

	set rsCheck = Server.CreateObject("ADODB.Connection")
	rsCheck.Open MM_bodyartforms_sql_STRING

	sqlString = "SELECT j.ProductID, j.title, j.description, j.picture, j.onsale, j.salediscount, CAST(j.ProductNotes AS VARCHAR(MAX)) as ProductNotes, j.type, MAX(pd.DateLastPurchased) as 'LastPurchaseDate',"
	sqlString = sqlString & "MIN(pd.DateLastPurchased) as 'OldestPurchaseDate' FROM jewelry j inner join dbo.ProductDetails pd "
	sqlString = sqlString & "on j.ProductID = pd.ProductID WHERE j.Active = 1 and j.jewelry not like '%save%' and j.title not like '%pre-order%' "
	sqlString = sqlString & "group by j.ProductID, j.title, j.description, j.picture, j.onsale, j.salediscount, CAST(j.ProductNotes AS VARCHAR(MAX)), j.type"
	sqlString = sqlString & " HAVING (MAX(pd.DateLastPurchased) <= DATEADD(month, -" & rangefilter & ", getdate())) ORDER BY MAX(pd.DateLastPurchased)"
	sqlString = sqlString & "  OFFSET " & offset & " ROWS FETCH NEXT " & pageSize & " ROWS ONLY"
	QueryToJSON(rsCheck, sqlString).Flush

	
end if
%>