<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET BackorderReview = 'Y', notes = ? WHERE OrderDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("notes",200,1,50, request.form("notes") ))
objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,20, request.form("orderdetailid") ))
objCmd.Execute  

DataConn.Close()
%>