<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_Sliders SET active=? WHERE sliderID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("active",3,1,10,Request.Form("active")))
objCmd.Parameters.Append(objCmd.CreateParameter("sliderID",3,1,10,Request.Form("id")))
Set rs_getImage_Filename = objCmd.Execute()

DataConn.Close()
%>