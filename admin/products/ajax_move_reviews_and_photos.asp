<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_PhotoGallery SET orig_productid = ProductID, orig_detailid = DetailID, ProductID = ? WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("param1",3,1,10,Request.Form("new_productid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("param2",3,1,10,Request.Form("old_productid")))
	objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBLReviews SET orig_productid = ProductID, orig_detailid = ISNULL(DetailID,0), ProductID = ? WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("param1",3,1,10,Request.Form("new_productid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("param2",3,1,10,Request.Form("old_productid")))
	objCmd.Execute()

DataConn.Close()
%>