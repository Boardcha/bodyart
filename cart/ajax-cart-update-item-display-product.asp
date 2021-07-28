<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
' ======= Get product info ===========================
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT picture, title FROM jewelry WHERE ProductID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,20, request.form("productid")))
Set rsGetProduct = objCmd.Execute()

if NOT rsGetProduct.eof then
%>
<img class="float-left mr-2 mb-1"  src="https://s3.amazonaws.com/bodyartforms-products/<%= rsGetProduct.Fields.Item("picture").Value %>" alt="Product photo">
<%= rsGetProduct.Fields.Item("title").Value %>

<%
end if ' record is found

DataConn.Close()
Set DataConn = Nothing
%>