<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
If request.form("type") <> "" OR request.querystring("clean") = "yes" then
	session("type") = request.form("type")
	session("detail_id") = ""
	session("bin") = ""
	session("pic") = ""
	session("description") = ""
	session("db_bin") = ""
	session("finished") = ""
End if

'==== If a period is found on the scanning label then parse it out to get the detail id. Otherwise default to a regular label that doesn't have the format in the barcode of purchaseorderid.detailid
if instr(request.form("Item"), ".") then
	item_array = split(request.form("Item"), ".")
	var_item_id = item_array(1)
else
	var_item_id = request.form("Item")
end if

If var_item_id <> "" AND request.form("needid") = "" AND request.form("needbin") = "" then
	session("detail_id") = var_item_id
end if

If var_item_id <> "" AND request.form("needid") = 1 then
	session("detail_id") = var_item_id
	
	Set cmd = Server.CreateObject ("ADODB.Command")
	cmd.ActiveConnection = MM_bodyartforms_sql_STRING
	cmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.ProductID, jewelry.picture, jewelry.title, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1, ProductDetails.BinNumber_Detail FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE (ProductDetails.ProductDetailID = " & session("detail_id") & ")"
	Set GetPic = cmd.Execute
	
If NOT GetPic.EOF then
	session("pic") = GetPic.Fields.Item("picture").Value
	session("description") = GetPic.Fields.Item("gauge").Value & " " & GetPic.Fields.Item("title").Value
	session("productid") = GetPic.Fields.Item("ProductID").Value
	
	If GetPic.Fields.Item("BinNumber_Detail").Value <> 0 then
		session("db_bin") = GetPic.Fields.Item("BinNumber_Detail").Value ' bin # already stored in the db
	else
		session("db_bin") = ""
	End if
	
End if 

	Set GetPic = Nothing
End if

If var_item_id <> "" AND request.form("needbin") = 1 then
	session("bin") = var_item_id
End if

If session("detail_id") <> "" AND session("type") <> "Limited" AND session("type") <> "Re-assign limited" then

	  If session("type") = "Limited" or session("type") = "Re-assign limited" then
		  setType = 0
	  End if
	  If session("type") = "Large" then
		  setType = 1
	  End if
	  If session("type") = "Party" then
		  setType = 3
	  End if
	  If session("type") = "Clothing" then
		  setType = 4
	  End if
	  If session("type") = "Pegboard" then
		  setType = 6
	  End if
	  If session("type") = "Regular" then
		  setType = 0
	  End if
	  If session("type") = "Vinyl" then
		  setType = 7
	  End if
	 If session("type") = "A" then
		  setType = 8
	  End if
	    If session("type") = "B" then
		  setType = 9
	  End if
	  
	  
		  set cmd = Server.CreateObject("ADODB.Command")
		  cmd.ActiveConnection = MM_bodyartforms_sql_STRING
		  cmd.CommandText = "UPDATE ProductDetails SET DetailCode = " & setType & ", active = 1 WHERE ProductDetailID = " & session("detail_id")
		  cmd.Execute()
		  Set cmd = Nothing
		  
		  set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = MM_bodyartforms_sql_STRING  
			objCmd.CommandText = "UPDATE jewelry SET active = 1 FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = " & session("detail_id")
			objCmd.Execute()
			Set objCmd = Nothing

			'Write info to edits log	
			set objCmd = Server.CreateObject("ADODB.Command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, product_id, detail_id, description, edit_date) VALUES (?," & session("productid") & "," & session("detail_id") & ",'Scanned into " & session("type")  & " BIN " & session("bin") & "','" & now() & "')"
			objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
			objCmd.Execute()
			Set objCmd = Nothing
		  
		  session("finished") = "yes"
		 ' response.redirect "barcode_convert.asp"

end if	

'response.write "Detail id -" & session("detail_id") & "-<br/>"
'response.write "bin -" & session("bin") & "-<br/>"
'response.write "finished -" & session("finished") & "-<br/>"
'response.write "db_bin -" & session("db_bin") & "-<br/>"
'response.write "Form_item -" & request.form("Item") & "-<br/>"
'response.write "Form_needID -" & request.form("needid") & "-<br/>"
'response.write "Form_needBIN -" & request.form("needbin") & "-<br/>"

If session("detail_id") <> "" AND session("bin") <> "" AND session("finished") <> "yes"  then	

		  set cmd = Server.CreateObject("ADODB.Command")
		  cmd.ActiveConnection = MM_bodyartforms_sql_STRING
		  cmd.CommandText = "UPDATE ProductDetails SET BinNumber_Detail = " & session("bin") & ", DetailCode = 0, active = 1 WHERE ProductDetailID = " & session("detail_id")
		  cmd.Execute()
		  Set cmd = Nothing
		  
		  set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = MM_bodyartforms_sql_STRING  
			objCmd.CommandText = "UPDATE jewelry SET active = 1 FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = " & session("detail_id")
			objCmd.Execute()
			Set objCmd = Nothing

			'Write info to edits log	
			set objCmd = Server.CreateObject("ADODB.Command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, product_id, detail_id, description, edit_date) VALUES (?," & session("productid") & "," & session("detail_id") & ",'Scanned into " & session("type")  & " BIN " & session("bin") & "','" & now() & "')"
			objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
			objCmd.Execute()
			Set objCmd = Nothing
		  
		  session("finished") = "yes"
		 ' response.redirect "barcode_convert.asp"

end if		 
	  


%>
<html>
<head>
<title>Scan item to section</title>
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
<script src="https://use.fortawesome.com/dc98f184.js"></script>
<link href="/CSS/baf.min.css?v=092519" rel="stylesheet" type="text/css" />
</head>
<body>
<!--#include file="includes/scanners-header.asp" -->
<div class="p-2">
<form id="FRM_ItemScan" name="FRM_ItemScan" method="post" action="barcode_convert.asp">

    <% If session("type") = "" OR request.querystring("clean") = "yes" then %>
<select class="form-control form-control-sm mb-2" name="type" id="type">
  <option value="Limited" selected>Limited</option>
  <option value="Case 1">Case 1</option>
  <option value="Case 2">Case 2</option>
  <option value="Large">Large</option>
  <option value="Party">Party</option>
  <option value="Pegboard">Pegboard</option>
  <option value="Clothing">Clothing</option>
  <option value="Regular">Regular</option>
  <option value="Vinyl">Vinyl</option>
  <option value="A">A</option>
  <option value="B">B</option>
  <option value="Re-assign limited">Re-assign limited</option>
</select>

    <input class="btn btn-primary" type="submit" name="button" id="button" value="Submit">
   
<% else ' if session.type is not empty %>
<div class="mb-2">
	<span class="font-weight-bold mr-2">Scanning into: <%= session("type") %></span><a class="btn btn-sm btn-outline-secondary" href="barcode_convert.asp?clean=yes">Start over</a>
</div>
 <%' Only process page below if there is NO BIN # assigned, otherwise display it
If session("db_bin") <> "" AND session("type") <> "Re-assign limited" then
session("finished") = "yes"  %>
<div class="alert alert-danger h6">ALREADY IN BIN # <%= session("db_bin") %>
</div>
<% end if %>
 <% if session("type") <> "Limited" AND session("type") <> "Re-assign limited" then
 	var_placeholder = "ITEM" %>
  <input name="needid" type="hidden" id="needid" value="1">
<% else ' if limited scanning
If session("detail_id") = "" OR (session("detail_id") <> "" AND session("finished") = "yes") OR (session("detail_ID") <> "" AND session("db_bin") <> "") then
	var_placeholder = "ITEM"
%>
<% if (session("bin") = "" AND session("type") = "Re-assign limited") then %>
<% else %>
	<input name="needid" type="hidden" id="needid" value="1">
<% 
end if
end if %>

<%
If (session("bin") = "" AND session("detail_id") <> "" AND session("db_bin") = "") OR (session("bin") = "" AND session("detail_id") <> ""AND session("type") = "Re-assign limited") then
	var_placeholder = "BIN"
%>
  <input name="needbin" type="hidden" id="needbin" value="1">
<% 
end if %>
<% end if ' if limited scanning %>
	<input class="form-control mt-1 font-weight-bold" name="Item" type="text" id="Item" placeholder="SCAN <%= var_placeholder %> #" autofocus />
<%
If session("finished") = "yes" then
	if session("detail_id") <> "" AND session("type") <> "Limited" AND session("type") <> "Re-assign limited" then
%>
	<div class="alert alert-success">Done</div>
<%	end if
	if session("bin") <> "" then
%>
	<div class="alert alert-success h5 mt-2">Scanned into BIN # <%= session("bin") %></div>
<%	end if

end if
%>
<% if session("detail_id") <> "" then %>
<div class="card bg-light mt-2">
	<div class="card-header p-2 font-weight-bold">
		<%= session("description") %>
		<div class="small">
			Detail ID #: <%= session("detail_id") %>
		</div>
	</div>
	<div class="card-body p-2">
		  <img src="http://bodyartforms-products.bodyartforms.com/<%= session("pic") %>" width="150" height="150">
	</div>
  </div>
<% end if %>

  <% end if  ' session("detail_id") <> ""  %>

</form>
</div>
</body>
</html>
<%
If session("finished") = "yes" then
	
	session("detail_id") = ""
	session("bin") = ""
	session("pic") = ""
	session("description") = ""
	session("finished") = ""
	session("db_bin") = ""
	
	
end if
%>