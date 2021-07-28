<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/cart/generate_guest_id.asp"-->
<% 
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE tbl_guest_users SET guest_email = ? WHERE guest_unique_id = " & var_cart_userid
    objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100,request.form("email")))
    objCmd.Execute()

DataConn.Close()
Set DataConn = Nothing
%>