<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"
%>
<html>
<head>
<title>Edit gallery photo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h4>          
	  Move customer photo to a different product
	</h4> 

 <% IF request.querystring("Update") = "Yes" then

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE dbo.TBL_PhotoGallery SET ProductID = ? WHERE PhotoID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("param1",3,1,10,Request.Form("ProductID")))
		objCmd.Parameters.Append(objCmd.CreateParameter("param2",3,1,10,Request.Form("PhotoID")))
		objCmd.Execute()
%>
<span class="text-success font-weight-bold">Photo updated</span>
<%
end if 
%>         
<form class="form-inline" ACTION="Gallery_EditPhoto.asp?Edit=Yes" METHOD="POST" name="FRM_AddCompany" id="FRM_AddCompany">
Search photo ID
  <input class="ml-2 mr-4 form-control form-control-sm" name="PhotoID" type="text" id="PhotoID" size="10"> 
  <button class="btn btn-sm btn-secondary" type="submit" name="Submit2">Submit</button>

</form>
<% if request.querystring("Edit") = "Yes" then %>
<% 
Dim rsGetPhoto
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT PhotoID, ProductID, filename, customerID, name, email, DateSubmitted FROM dbo.TBL_PhotoGallery WHERE PhotoID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("param1",3,1,10,Request.Form("PhotoID")))
		Set rsGetPhoto = objCmd.Execute()

if rsGetPhoto.Fields.Item("DateSubmitted").Value < "2/20/2011" then
DomainLink = "bodyartforms-gallery.bodyartforms.com" ' AMAZON S3
'DomainLink = "216.128.23.22/gallery/uploads"
Else
DomainLink = "www.bodyartforms.com/gallery/uploads"
End if
%>
<form class="form-inline" ACTION="Gallery_EditPhoto.asp?Update=Yes" METHOD="POST" name="FRM_EditPhoto" id="FRM_EditPhoto">
<% if NOT rsGetPhoto.BOF AND NOT rsGetPhoto.EOF then %>
<hr size="1">
<p><img src="http://<%= DomainLink %>/thumb_<%= Replace(rsGetPhoto.Fields.Item("filename").Value, ".JPG", ".jpg") %>"/></p>
	<br/><br/>
	Customer ID # <%= rsGetPhoto.Fields.Item("customerID").Value %><br/>
	<%= rsGetPhoto.Fields.Item("name").Value %></br>
	<%= rsGetPhoto.Fields.Item("email").Value %>
  <p>Move to product ID #
    <input class="ml-2 mr-4 form-control form-control-sm" name="ProductID" type="text" id="ProductID" size="10">
    <button class="btn btn-sm btn-secondary" type="submit" name="Submit">Update</button>
    <input name="PhotoID" type="hidden" id="PhotoID" value="<%= rsGetPhoto.Fields.Item("PhotoID").Value %>">
	  </p>
<% else %>
  Photo not found
  <% end if %>
</form>
<% end if %>
 
       
</div>
</body>
</html>
<%
DataConn.Close()
Set DataConn = Nothing
set objCmd = Nothing
%>