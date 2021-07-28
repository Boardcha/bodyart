<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
if request.form("add") = "add" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_WishlistCategories (WishlistName, Wishlist_CustomerID) VALUES (?, ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("category",200,1,50, request.form("category")))
	objCmd.Parameters.Append(objCmd.CreateParameter("custID",3,1,12, CustID_Cookie))
	objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_WishlistCategories WHERE Wishlist_CustomerID = ? ORDER BY WishlistID DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("custID",3,1,12, CustID_Cookie))
	set rsGetNewCategory = objCmd.Execute()
	
	
	if NOT rsGetNewCategory.eof then
%>	
{
	"category_id":"<%= rsGetNewCategory.Fields.Item("WishlistID").Value %>"
}
<%	
	end if

end if ' if add ------------------------------------

if request.form("retrieve") = "retrieve" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_WishlistCategories WHERE Wishlist_CustomerID = ? ORDER BY WishlistName ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("custID",3,1,12, CustID_Cookie))
	set rsGetCategories = objCmd.Execute()

if not rsGetCategories.eof then
%>
<form>
	<input type="hidden" name="update-categories" value="yes" />
<h6 class="mt-3">Current Lists:</h6>
<%	
	while not rsGetCategories.eof
%>
<div class="my-2" id="category-<%= rsGetCategories.Fields.Item("WishlistID").Value %>">
	<input class="form-control form-control-sm w-auto d-inline" type="text" name="update-category-name" maxlength="50" id="update-<%= rsGetCategories.Fields.Item("WishlistID").Value %>" value="<%= rsGetCategories.Fields.Item("WishlistName").Value %>">
	<button class="btn btn-sm btn-outline-success ml-2 btn-update-category" data-id="<%= rsGetCategories.Fields.Item("WishlistID").Value %>" type="button">Update</button>
	<button class="btn btn-sm btn-outline-danger ml-2 delete-category" data-id="<%= rsGetCategories.Fields.Item("WishlistID").Value %>" type="button"><i class="fa fa-trash-alt"></i></button>
</div>
<%	
	rsGetCategories.movenext()
	wend
%>
</form>
<%
end if 'if not rsGetCategories.eof
end if ' retrieve records for updating --------------

' Response.Write "New value " & Request.Form("new_name")
if request.form("update") = "yes" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_WishlistCategories SET WishlistName = ? WHERE WishlistID = ? AND Wishlist_CustomerID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("new_name",200,1,50, request.form("new_name")))
	objCmd.Parameters.Append(objCmd.CreateParameter("item_id",3,1,12, request.form("category_id")))
	objCmd.Parameters.Append(objCmd.CreateParameter("custID",3,1,12, CustID_Cookie))
	objCmd.Execute()

end if ' update categories --------------------------


if request.form("delete") = "yes" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_WishlistCategories WHERE WishlistID = ? AND Wishlist_CustomerID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("category_id",3,1,12, request.form("category_id")))
	objCmd.Parameters.Append(objCmd.CreateParameter("custID",3,1,12, CustID_Cookie))
	objCmd.Execute()

end if ' delete category ----------------------------


DataConn.Close()
Set DataConn = Nothing
%>
