<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
var_cleaned_review = replace(replace(replace(replace(request.form("review"), "’", "'"), "‘", "'"), "”", """"), "“", """")

response.write "rating: " & request.form("rating") & "<br>"
response.write "review: " & request.form("review") & "<br>"
response.write "detail id: " & request.form("id") & "<br>"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT  TOP (1) OrderDetailID, DetailID, qty, ProductID, InvoiceID FROM TBL_OrderSummary WHERE  OrderDetailID = ?"
objCmd.NamedParameters = True
objCmd.Parameters.Append(objCmd.CreateParameter("@ProductID",3,1,10, request.form("id")))
Set rsGetProduct = objCmd.Execute()

' Pull the customer information from a cookie
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (1) customer_first, email  FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10, CustID_Cookie))
Set rsGetUser = objCmd.Execute()

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT  TOP (1) ReviewID FROM TBLReviews WHERE ReviewOrderDetailID = ?"
objCmd.NamedParameters = True
objCmd.Parameters.Append(objCmd.CreateParameter("@ProductID",3,1,10, request.form("id")))
Set rsCheckDupe = objCmd.Execute()

if rsCheckDupe.eof then ' if there's no review already found for that order detail # then insert a new row

	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "INSERT INTO TBLReviews (name, email, review, review_rating, date_submitted, status, ProductID, customer_ID, ReviewOrderDetailID, DetailID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
	  objCmd.Parameters.Append(objCmd.CreateParameter("@Name",200,1,50, rsGetUser.Fields.Item("customer_first").Value))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@Email",200,1,75, rsGetUser.Fields.Item("email").Value))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@Review",200,1,2000, var_cleaned_review))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@Ratingg",3,1,12, request.form("rating")))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@DateSubmitted",200,1,20, date()))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@status",200,1,20, "pending"))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@ProductID",3,1,12, rsGetProduct.Fields.Item("ProductID").Value))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@CustomerID",3,1,20, CustID_Cookie))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@ReviewOrderDetailID",3,1,20, rsGetProduct.Fields.Item("OrderDetailID").Value))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@DetailID",3,1,20, rsGetProduct.Fields.Item("DetailID").Value))
	  objCmd.Execute()
	  
	  ' set to status reviewed Y in database in orders table
	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = " UPDATE TBL_OrderSummary SET ProductReviewed = 'Y' WHERE OrderDetailID = ?"
	  objCmd.Parameters.Append(objCmd.CreateParameter("@OrderDetailID",3,1,20, rsGetProduct.Fields.Item("OrderDetailID").Value))
	  objCmd.Execute()

else ' if a review already exists then just update the row

	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "UPDATE TBLReviews SET review = ?, status = ? WHERE ReviewID = ?"
	  objCmd.Parameters.Append(objCmd.CreateParameter("@Review",200,1,2000, var_cleaned_review))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@status",200,1,20, "pending"))
	  objCmd.Parameters.Append(objCmd.CreateParameter("@ReviewOrderDetailID",3,1,20, rsCheckDupe.Fields.Item("ReviewID").Value))
	  objCmd.Execute()

end if ' if review already exists or not

DataConn.Close()
Set DataConn = Nothing
%>