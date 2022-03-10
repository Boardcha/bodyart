<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
var_bo_reason = Request.Form("bo_reason")
If var_bo_reason <> "" Then param_bo_reason = ", reason_for_backorder = '" + var_bo_reason + "'"

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET BackorderReview = 'Y', notes = ? " & param_bo_reason & " WHERE OrderDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("notes",200,1,50, request.form("notes") ))
objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,20, request.form("orderdetailid") ))
objCmd.Execute  
Response.Write "UPDATE TBL_OrderSummary SET BackorderReview = 'Y', notes = ? " & param_bo_reason & " WHERE OrderDetailID = ?" 
DataConn.Close()
%>