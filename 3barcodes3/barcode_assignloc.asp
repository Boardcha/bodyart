<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
If request.form("type") <> "" OR request.querystring("clean") = "yes" then
	session("type") = request.form("type")
	session("detail_id") = ""
	session("bin") = ""
	session("pic") = ""
	session("description") = ""
	session("finished") = ""
	session("location") = ""
End if

If request.form("location") <> "" then
	session("location") = request.form("location")
else
	session("location") = session("location")
end if

If request.form("Item") <> "" AND request.form("needid") = 1 then
	session("detail_id") = request.form("Item")
	
	Set cmd = Server.CreateObject ("ADODB.Command")
	cmd.ActiveConnection = MM_bodyartforms_sql_STRING
	cmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.picture, jewelry.title, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1, ProductDetails.BinNumber_Detail FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE (ProductDetails.ProductDetailID = " & session("detail_id") & ")"
	Set GetPic = cmd.Execute
	
If NOT GetPic.EOF then
	session("pic") = GetPic.Fields.Item("picture").Value
	session("description") = GetPic.Fields.Item("gauge").Value & " " & GetPic.Fields.Item("title").Value
	
	
End if 

	Set GetPic = Nothing
End if

'If request.form("Item") <> "" AND request.form("needbin") = 1 then
'	session("bin") = request.form("Item")
'End if

If session("detail_id") <> "" then

	If session("type") = 0 then
	setType = "Limited"
	elseif session("type") = 1 then
	setType = "Large"
	elseif session("type") = 3 then
	setType = "Party"
	elseif session("type") = 4 then
	setType = "Clothing"
	elseif session("type") = 5 then
	setType = "Pegboard"
	elseif session("type") = 0 then
	setType = "Regular"
	elseif session("type") = 7 then
	setType = "Vinyl"
	elseif session("type") = 8 then
	setType = "A"
	elseif session("type") = 9 then
	setType = "B"
	elseif session("type") = 10 then
	setType = "C"
	elseif session("type") = 11 then
	setType = "D"
	elseif session("type") = 12 then
	setType = "E"
	elseif session("type") = 13 then
	setType = "F"
	elseif session("type") = 14 then
	setType = "G"
	elseif session("type") = 15 then
	setType = "H"
	elseif session("type") = 16 then
	setType = "I"
	elseif session("type") = 17 then
	setType = "J"
	elseif session("type") = 18 then
	setType = "K"
	elseif session("type") = 19 then
	setType = "L"
	elseif session("type") = 20 then
	setType = "M"
	elseif session("type") = 21 then
	setType = "N"
	elseif session("type") = 22 then
	setType = "O"
	elseif session("type") = 23 then
	setType = "P"
	elseif session("type") = 24 then
	setType = "Q"
	elseif session("type") = 25 then
	setType = "R"
	elseif session("type") = 26 then
	setType = "S"
	elseif session("type") = 27 then
	setType = "T"
	elseif session("type") = 28 then
	setType = "U"
	elseif session("type") = 29 then
	setType = "V"
	elseif session("type") = 30 then
	setType = "W"
	elseif session("type") = 31 then
	setType = "X"
	elseif session("type") = 32 then
	setType = "Y"
	elseif session("type") = 33 then
	setType = "Z"
	End if
 
	  
		  set cmd = Server.CreateObject("ADODB.Command")
		  cmd.ActiveConnection = MM_bodyartforms_sql_STRING
		  cmd.CommandText = "UPDATE ProductDetails SET DetailCode = " & session("type") & ", location = " & session("location") & " WHERE ProductDetailID = " & session("detail_id")
		  cmd.Execute()
		  Set cmd = Nothing

		  'Write info to edits log	
		  set objCmd = Server.CreateObject("ADODB.Command")
		  objCmd.ActiveConnection = DataConn
		  objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, product_id, detail_id, description, edit_date) VALUES (6," & session("productid") & "," & session("detail_id") & ",'Scanned into section " & setType  & " Location " & session("location") & "','" & now() & "')"
		  objCmd.Execute()
		  Set objCmd = Nothing
		  
		  session("location") = session("location") + 1
		  session("finished") = "yes"

end if		 
%>
<html>
<head>
<title>Scan item to section</title>
<style type="text/css">
<!--
body {
		-webkit-text-size-adjust:none;
	  	font-family: Helvetica, Arial, Verdana, sans-serif;
	  	font-size: 15px;
	  	color: black;
	  }
	  
 
.ShowItems {
	display: inline;
	font-size: 15px;
	margin: 0px;
	}
	
.notes {
	   color: blue;
	   font-size: 15px;
	   font-weight: bold;
	   }	
	   
.accent {
		font-weight: bold;
		font-size: 16px;
			}   

.stillneed {
		   color: #3300FF;
		   font-weight: bold;
		   font-size: 20px;
		   }
.alert {
	   color: #CC0000;
	   font-weight: bold;
	   font-size: 25px;
	   }
-->
</style>
</head>
<body class="materialText">
<form id="FRM_ItemScan" name="FRM_ItemScan" method="post" action="barcode_assignloc.asp">
  <p>
    <% If session("type") = "" OR request.querystring("clean") = "yes" then %>
<select name="type" id="type">
<option value="8">A</option>
<option value="9">B</option>
  <option value="10">C</option>
  <option value="11">D</option>
  <option value="12">E</option>
<option value="13">F</option>
<option value="14">G</option>
<option value="15">H</option>
<option value="16">I</option>
<option value="17">J</option>
<option value="18">K</option>
<option value="19">L</option>
<option value="20">M</option>
<option value="21">N</option>
<option value="22">O</option>
<option value="23">P</option>
<option value="24">Q</option>
<option value="25">R</option>
<option value="26">S</option>
<option value="27">T</option>
<option value="28">U</option>
<option value="29">V</option>
<option value="30">W</option>
<option value="31">X</option>
<option value="32">Y</option>
<option value="33">Z</option>
<option value="3">Extra large</option>
<option value="4">Clothing</option>
<option value="7">Vinyl</option>
</select>
 Starting location # 
 <input name="location" type="text" id="location" size="10" />
  </p>
  <p>
    <input type="submit" name="button" id="button" value="Submit">
    <br />
<% else ' if session.type is not empty %>
  <span class="stillneed">Scan the DETAIL ID #</span><br><br>
  <input name="needid" type="hidden" id="needid" value="1">
  <span style="font-size: 14px;"><strong>Scanning into: <%= session("type") %>, Location: <%= session("location") %>&nbsp;&nbsp;&nbsp;</strong></span><a href="barcode_assignloc.asp?clean=yes">Start over</a>
<%
If session("finished") = "yes" then
	if session("detail_id") <> "" AND session("type") <> "Limited" then
%>
	<br>
<span class="alert">Done</span>
<%	end if
end if
%>
<% if session("detail_id") <> "" then %>
<br />
<br>
Detail ID #: <%= session("detail_id") %><br>
<%= session("description") %><br>
<img src="http://bodyartforms-products.bodyartforms.com/<%= session("pic") %>" width="80" height="80">
<% end if %>
<input name="Item" type="text" id="Item" size="10" />
  </p>
  <% end if %>
  </p>
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
	
end if
%>