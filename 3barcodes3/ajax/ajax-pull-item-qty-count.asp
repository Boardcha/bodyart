<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->

{
<%
status = request.form("status")


if status = "update" then

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE ProductDetails SET qty_counted_discontinued = ?, item_pulled = 1 WHERE ProductDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("qty_counted",3,1,15, request.form("qty_counted") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, request.form("detailid") ))
    set rsGetDetails = objCmd.Execute()

    ' ---- Check to see if there are any details left and if not, then send back json response
    ' --- pull details
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT item_pulled FROM ProductDetails WHERE ProductID = ? AND item_pulled = 0"
    objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20, request.form("productid")  ))
    set rsGetDetails = objCmd.Execute()
    
    if rsGetDetails.eof then
    %>
        "status":"complete",
        "action":"update",
        "product_id": "<%= request.form("productid") %>"
<%  else %>
        "status":"incomplete",
        "action":"update",
        "product_id": "<%= request.form("productid") %>"
<%
    end if 

else

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE ProductDetails SET qty_counted_discontinued = ?, item_pulled = 0 WHERE ProductDetailID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("qty_counted",3,1,15, request.form("qty_counted") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, request.form("detailid") ))
    set rsGetDetails = objCmd.Execute()
%>
        "status":"incomplete",
        "action":"clear"
<%
end if
%>
}