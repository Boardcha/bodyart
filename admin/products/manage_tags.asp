<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetTags
Dim rsGetTags_numRows
Set rsGetTags = Server.CreateObject("ADODB.Recordset")
rsGetTags.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetTags.Source = "SELECT * FROM dbo.TBL_Product_Tags ORDER BY Tag ASC"
rsGetTags.CursorLocation = 3 'adUseClient
rsGetTags.LockType = 1 'Read-only records
rsGetTags.Open()
%>
<html>
<head>
<title>Manage Tags</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/CSS/baf.min.css?v=040220" id="lightmode" rel="stylesheet" type="text/css">
<script src="https://use.fortawesome.com/dc98f184.js"></script>
</head>
<body>
<div class="p-3">
<form name="FRM_AddNewTag" id="FRM_AddNewTag">
	<h5>Add a New Tag</h5>
	<table class="table table-sm table-borderless w-25">
		<tr> 
		  <td>Tag name</td>
		  <td> <input name="tag" type="text" maxlength ="50" class="form-control form-control-sm" id="tag"></td>
		</tr>
		<tr>
		  <td>&nbsp;</td>
		  <td><input class="btn btn-purple" type="button" name="add_new_tag" id="add_new_tag" value="Add">
		</tr>
	</table>
</form>
<h4>Tags</h4>
<table class="table table-striped table-borderless table-hover w-50">
<% While (NOT rsGetTags.EOF) %>
    <tr>
      <td>
          <a class="btn btn-sm btn-danger mr-4 delete-tag" data-tag-id="<%=(rsGetTags.Fields.Item("TagID").Value)%>"><i class="fa fa-trash-alt text-white"></i></a>
          <%=(rsGetTags.Fields.Item("Tag").Value)%>
	  </td>
    </tr>
<% 
rsGetTags.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsGetTags.Close()
Set rsGetTags = Nothing
%>
