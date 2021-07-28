<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
arrIDs = Split(Request.Form("data"), ",")
For each id in arrIDs
	new_order = new_order + 1
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_Sliders SET show_up_order=? WHERE sliderID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("sliderID",3,1,10,new_order))
	objCmd.Parameters.Append(objCmd.CreateParameter("sliderID",3,1,10,id))
	Set rs_getImage_Filename = objCmd.Execute()
Next 

DataConn.Close()
%>