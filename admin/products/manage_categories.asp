<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetCategories
Dim rsGetCategories_numRows
Set rsGetCategories = Server.CreateObject("ADODB.Recordset")
rsGetCategories.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCategories.Source = "SELECT * FROM dbo.TBL_Categories ORDER BY category_name ASC"
rsGetCategories.CursorLocation = 3 'adUseClient
rsGetCategories.LockType = 1 'Read-only records
rsGetCategories.Open()
%>
<html>
<head>
<title>Manage Categories</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/CSS/baf.min.css?v=040220" id="lightmode" rel="stylesheet" type="text/css">
<script src="https://use.fortawesome.com/dc98f184.js"></script>
</head>
<body>
<div class="p-3">
<form name="FRM_AddNewCategory" id="FRM_AddNewCategory">
	<h5>Add a New Category</h5>
	<table class="table table-sm table-borderless w-50">
		<tr> 
		  <td>Admin friendly name</td>
		  <td> <input name="category_name" type="text" maxlength ="100" class="form-control form-control-sm" id="category_name"></td>
		</tr>
		<tr> 
		  <td>Database search tag (no spaces)</td>
		  <td> <input name="category_tag" type="text" maxlength ="50" class="form-control form-control-sm" id="category_tag"></td>
		</tr>		
		<tr>
		  <td>&nbsp;</td>
		  <td><input class="btn btn-purple" type="button" name="add_new_category" id="add_new_category" value="Add">
		</tr>
	</table>
</form>
<h4>Categories</h4>
<table class="table table-striped table-borderless table-hover w-75">
	<thead class="thead-dark">
		<tr>
		<th>Admin friendly name</th>
		<th>Database search tag</th>
	</tr>
	</thead>
<<<<<<< HEAD
=======

>>>>>>> c02b5fc9e9a062d35f985c4be9732f2184e9381c
<% While (NOT rsGetCategories.EOF) %>
    <tr>
      <td>
          <a class="btn btn-sm btn-danger mr-4 delete-category" data-category-id="<%=(rsGetCategories.Fields.Item("category_id").Value)%>"><i class="fa fa-trash-alt text-white"></i></a>
          <%=(rsGetCategories.Fields.Item("category_name").Value)%>
	  </td>
	  <td>
		<%=(rsGetCategories.Fields.Item("category_tag").Value)%>
	</td>
    </tr>
<% 
rsGetCategories.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsGetCategories.Close()
Set rsGetCategories = Nothing
%>
