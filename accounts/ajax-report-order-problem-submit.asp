<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
varDescription = Replace(Request.form("description"), "'", "")

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_OrderSummary SET item_problem = ?, ErrorReportDate = ?, ErrorDescription = ?, ErrorOnReview = 1, ErrorQtyMissing = ? WHERE OrderDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("Problem",200,1,30,request.form("status")))
	objCmd.Parameters.Append(objCmd.CreateParameter("ReportDate",200,1,12,date()))
	objCmd.Parameters.Append(objCmd.CreateParameter("Description",200,1,500,varDescription))
	objCmd.Parameters.Append(objCmd.CreateParameter("QtyMissing",3,1,2,Request.form("qty-missing")))
	objCmd.Parameters.Append(objCmd.CreateParameter("DetailID",3,1,10,Request.form("report-item")))
	objCmd.Execute()

DataConn.Close()
Set DataConn = Nothing
%>
