<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET ErrorQtyMissing = ? WHERE OrderDetailID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("value",3,1,5, request.form("value")))
objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,20, request.form("id")))
set rsGetInvoice = objCmd.Execute()


DataConn.Close()
%>

