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

If request.form("Item") <> "" AND request.form("needid") = 1 then
	session("detail_id") = request.form("Item")
	
	Set cmd = Server.CreateObject ("ADODB.Command")
	cmd.ActiveConnection = MM_bodyartforms_sql_STRING
	cmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.picture, jewelry.title, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1, ProductDetails.BinNumber_Detail, ProductDetails.ProductID FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE (ProductDetails.ProductDetailID = " & session("detail_id") & ")"
	Set GetPic = cmd.Execute
	
If NOT GetPic.EOF then
	session("productid") = GetPic.Fields.Item("ProductID").Value
	session("pic") = GetPic.Fields.Item("picture").Value
	session("description") = GetPic.Fields.Item("gauge").Value & " " & GetPic.Fields.Item("title").Value
	
	If GetPic.Fields.Item("BinNumber_Detail").Value <> 0 then
		session("db_bin") = GetPic.Fields.Item("BinNumber_Detail").Value ' bin # already stored in the db
	else
		session("db_bin") = ""
	End if
	
End if 

	Set GetPic = Nothing
End if

If request.form("Item") <> "" AND request.form("needbin") = 1 then
	session("bin") = request.form("Item")
End if

If session("detail_id") <> "" AND session("bin") <> "" AND session("finished") <> "yes"  then

	  If session("type") = "Case 1" then
		  setType = 34
	  End if
	  If session("type") = "Case 2" then
		  setType = 35
	  End if	
	  If session("type") = "Case 3" then
		  setType = 36
	  End if	
	  If session("type") = "Case 4" then
		  setType = 37
  		End if

  'response.write "session detail id: " & session("detail_id")
  'response.write "<br/>bin #: " & session("bin")
  'response.write "<br/>setType: " & setType
	  
		  set objCmd = Server.CreateObject("ADODB.Command")
		  objCmd.ActiveConnection = DataConn
		  objCmd.CommandText = "UPDATE ProductDetails SET DetailCode = " & setType & ", BinNumber_Detail = " & session("bin") & " WHERE ProductDetailID = " & session("detail_id")
		  objCmd.Execute()
		  Set objCmd = Nothing



		  'Write info to edits log	
		  set objCmd = Server.CreateObject("ADODB.Command")
		  objCmd.ActiveConnection = DataConn
		  objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, product_id, detail_id, description, edit_date) VALUES (6," & session("productid") & "," & session("detail_id") & ",'Scanned into " & session("type")  & " SHELF " & session("bin") & "','" & now() & "')"
		  objCmd.Execute()
		  Set objCmd = Nothing
		  
		  session("finished") = "yes"
		 ' response.redirect "scan_front-cases.asp"

end if		 
	  


%>
<html>
<head>
<title>Scan item into front cases</title>
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
<link href="/CSS/baf.min.css?v=120318" rel="stylesheet" type="text/css" />
</head>
<body class="mx-2">
<form id="FRM_ItemScan" name="FRM_ItemScan" method="post" action="scan_front-cases.asp">
	<% If session("type") = "" OR request.querystring("clean") = "yes" then %>
	<div  class="form-inline mt-2">
		<select  class="form-control mr-3" name="type" id="type">
		<option value="Case 1">Case 1</option>
		<option value="Case 2">Case 2</option>
		<option value="Case 3">Case 3</option>
		<option value="Case 4">Case 4</option>
		</select>
	<input class="btn btn-primary" type="submit" name="button" id="button" value="Submit">
</div>
<% else ' if session.type is not empty
' Only process page below if there is NO BIN # assigned, otherwise display it
%>
<h6 class="d-inline-block pt-0 pr-3">Scanning into: <%= session("type") %></h6><a href="scan_front-cases.asp?clean=yes">Start over</a>
<%
If session("db_bin") <> "" and session("finished") <> "yes" then
	If session("type") = "Case 4" then
		ShelfBin = "ALREADY IN BIN"
	else
		ShelfBin = "ALREADY ON SHELF"
	End if
%>
<div class="alert alert-danger h5"><%= ShelfBin %> # <%= session("db_bin") %><br/>RESCAN AND PUT ITEM IN SAME LOCATION</div>
<% end if %>
<%
If session("finished") = "yes" then
	if session("detail_id") <> "" then
%>
<div class="alert alert-success">Scanned to shelf # <%= session("bin") %></div>
<%	end if 
end if
%>
<%
If session("detail_id") = "" OR (session("detail_id") <> "" AND session("finished") = "yes")  then
%>
<h5>Scan the ITEM</h5>
<input name="needid" type="hidden" id="needid" value="1">
<% end if %>

<%
If session("bin") = "" AND session("detail_id") <> "" then
	If session("type") = "Case 4" then
		ShelfBin = "BIN"
	else
		ShelfBin = "SHELF"
	End if
%>
<h5>Scan the <%= ShelfBin %></h5>
  <input name="needbin" type="hidden" id="needbin" value="1">
<% end if %>

<input  class="form-control mb-3" name="Item" type="text" id="Item" size="10" placeholder="Scan barcode" autofocus/>
<% if session("detail_id") <> "" then %>

	<img class="float-left" src="http://bodyartforms-products.bodyartforms.com/<%= session("pic") %>" width="80" height="80">
	<%= session("description") %>
<% end if %>
  <% end if %>
</form>
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