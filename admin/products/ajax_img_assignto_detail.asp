<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET img_id = ? WHERE ProductDetailID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("img_id",3,1,15,request.queryString("imgid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15,request.queryString("detailid")))
	Set rs_GetImgID = objCmd.Execute()

DataConn.Close()
%>