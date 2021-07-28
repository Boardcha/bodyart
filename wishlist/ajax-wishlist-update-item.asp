<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
' ------------------------------------------
' UPDATE FROM WISHLIST PAGE
' ------------------------------------------
if request.form("wishlist_id") <> "" then 

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE wishlist SET desired = ?, priority = ?, WishlistID = ?, comments = ? WHERE ID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("desired",3,1,12, request.form("desired")))
    objCmd.Parameters.Append(objCmd.CreateParameter("desired",3,1,5, request.form("priority")))
    objCmd.Parameters.Append(objCmd.CreateParameter("category",3,1,12, request.form("category")))
    objCmd.Parameters.Append(objCmd.CreateParameter("comments",200,1,300, request.form("comments")))
    objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,12, request.form("wishlist_id")))
    objCmd.Execute()

end if ' request.form("wishlist_id") <> ""


' ------------------------------------------
' UPDATE FROM PRODUCT DETAILS PAGE
' ------------------------------------------
if request.form("session") = "yes" then 


    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE wishlist SET priority = ?, WishlistID = ? WHERE ID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("desired",3,1,5, request.form("priority")))
    objCmd.Parameters.Append(objCmd.CreateParameter("category",3,1,12, request.form("category")))
    objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,12, session("last_wishlist_item")))
    objCmd.Execute()

end if ' product id not found

DataConn.Close()
Set DataConn = Nothing
%>