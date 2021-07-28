<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' ---- Assigns products to person logged in to pull discontinued items

response.write request.form("productid")

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE jewelry SET who_pulled = ? WHERE ProductID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("who",200,1,50, rsGetUser.Fields.Item("name").Value ))
    objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, request.form("productid") ))
    set rsGetDetails = objCmd.Execute()
%>

