<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	'========= DELETE ITEM FROM PURCHASE ORDER ====================================

        
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn  
        objCmd.CommandText = "UPDATE tbl_po_details SET po_received = 1 WHERE po_detailid = " & request.form("po_detailid") & " AND po_orderid = " & Session("po_id")
        objCmd.Execute()
        Set objCmd = Nothing
	
%>