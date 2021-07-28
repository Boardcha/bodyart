<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->

<%

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE tbl_carts SET cart_preorderNotes = ? WHERE cart_id = ? AND " & var_db_field & " = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("cart_preorderNotes",200,1,300, request.form("specs")))
objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10, request.form("cartid")))
objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, var_cart_userid))
objCmd.Execute()

DataConn.Close()
Set DataConn = Nothing
%>