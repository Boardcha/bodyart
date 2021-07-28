<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="emails/function-send-email.asp"-->

<%
orderdetailid = request.form("orderdetailid")
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, ProductID, DetailID, title, ProductDetail1, Gauge, stock_qty, OrderDetailID, email, customer_first, ISNULL(title,'') + ' ' + ISNULL(ProductDetail1,'') + ' ' + ISNULL(Gauge,'') + ' ' + ISNULL(Length,'') as 'item_description' FROM dbo.QRY_OrderDetails WHERE OrderDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,20, orderdetailid))
Set rsGetInfo = objCmd.Execute()

productdetailid = rsGetInfo.Fields.Item("DetailID").Value
var_customer_name = rsGetInfo.Fields.Item("customer_first").Value
var_customer_email = rsGetInfo.Fields.Item("email").Value
var_invoice_number = rsGetInfo.Fields.Item("InvoiceID").Value
var_item_description = Server.HTMLEncode(rsGetInfo.Fields.Item("item_description").Value)
var_bo_reason = Request.Form("bo_reason")

' Set item to backorder status (and not on review)
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET backorder = 1, BackorderReview = 'N' WHERE OrderDetailID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,20, orderdetailid))
objCmd.Execute()

' Update quantities on item according to selected drop-down
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE ProductDetails SET qty = ? WHERE ProductDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,20, request.form("bo_qty")))
objCmd.Parameters.Append(objCmd.CreateParameter("productdetailid",3,1,20,productdetailid))
objCmd.Execute()

mailer_type = "backorder"
%>
<!--#include virtual="emails/email_variables.asp"-->
<%

DataConn.Close()
Set rsGetInfo = Nothing
%>