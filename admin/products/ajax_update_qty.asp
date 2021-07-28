<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'duplicate product
if request.form("qty") <> "" then


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT qty FROM ProductDetails WHERE ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,request.form("detailid")))
	Set rs_getdetail = objCmd.Execute()
	
	'if teh qty is going from 0 to in stock, then update the restock date column
	if rs_getdetail.Fields.Item("qty").Value = 0 then
		var_restocked = ", DateRestocked = '" & FormatDateTime(now(),2) & "'"
	end if

	'Check to see if the database qty field matches the orig loaded form qty
	if cLng(rs_getdetail.Fields.Item("qty").Value) = cLng(request.form("origqty")) then
'		response.write "matched! & now overwrite the qty"
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET qty = " & request.form("qty") & " " & var_restocked & " WHERE ProductDetailID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,request.form("detailid")))
		objCmd.Execute() 
%>
{  
   "status":""
}
<%
	else ' if the original quantities dont match then find the difference and write proper qty
'		response.write "NOT matched! & now overwrite the qty<br/>DB qty: " & rs_getdetail.Fields.Item("qty").Value & " Form qty: " & request.form("qty") & " Form ORIG qty: " & request.form("origqty") & " Detail id: " &  request.form("detailid")
		
		if cLng(rs_getdetail.Fields.Item("qty").Value) > cLng(request.form("origqty")) then 
			'if database qty is GREATER than the orig qty saved in form then
		'	difference = cLng(rs_getdetail.Fields.Item("qty").Value) - cLng(request.form("qty"))
		'	new_qty = cLng(request.form("qty")) - cLng(difference)		
		else
			'if database qty is LESS than the orig qty saved in form t
			difference = cLng(request.form("qty")) - cLng(rs_getdetail.Fields.Item("qty").Value)
		end if
			
			new_qty = cLng(request.form("qty")) - cLng(difference)
		
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET qty = " & new_qty & " WHERE ProductDetailID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,request.form("detailid")))
		objCmd.Execute() 
%>
{  
   "status":"A customer has purchased this item. The qty amount you entered has been automatically reduced by <%= difference %>",
   "difference":"<%= new_qty %>"
}
<%	end if
	

	edits_description = "Updated DETAIL: Changed quantity from " & request.form("origqty") & " to " & request.form("qty")
		
	'Write ALL info to edits log table
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, product_id, detail_id, description, edit_date) VALUES (" & user_id & "," & request.form("id") & "," & request.form("detailid") & ",'" & edits_description & "','" & now() & "')"
	objCmd.Execute()

		
end if ' check that detailid has been passed to process page
DataConn.Close()
%>