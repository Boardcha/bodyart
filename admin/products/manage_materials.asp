<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetMaterials
Dim rsGetMaterials_numRows
Set rsGetMaterials = Server.CreateObject("ADODB.Recordset")
rsGetMaterials.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetMaterials.Source = "SELECT * FROM dbo.TBL_Materials ORDER BY material_name ASC"
rsGetMaterials.CursorLocation = 3 'adUseClient
rsGetMaterials.LockType = 1 'Read-only records
rsGetMaterials.Open()
%>
<html>
<head>
<title>Manage Materials</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/CSS/baf.min.css?v=040220" id="lightmode" rel="stylesheet" type="text/css">
<script src="https://use.fortawesome.com/dc98f184.js"></script>
</head>
<body>
<div class="p-3">
<form name="FRM_AddNewMaterial" id="FRM_AddNewMaterial">
	<h5>Add a New Material</h5>
	<table class="table table-sm table-borderless w-50">
		<tr> 
		  <td>Material name</td>
		  <td><input name="material" type="text" maxlength ="200" class="form-control form-control-sm" id="material"></td>
		</tr>
		<tr> 
		  <td>&nbsp;</td>
		  <td>
		  	<div class="custom-control custom-checkbox">
				<input name="iswearable" id="iswearable" type="checkbox" class="custom-control-input iswearable">
				<label class="custom-control-label" for="iswearable">Wearable</label>
			</div>
		  </td>
		</tr>		
		<tr>
		  <td>&nbsp;</td>
		  <td colspan="3"><input class="btn btn-purple mt-2" type="button" name="add_new_tag" id="add_new_material" value="Add">
		</tr>
	</table>
</form>

<table class="table table-striped table-borderless table-hover w-75">
	<thead class="thead-dark">
		<tr>
			<th>
				Material Name
			</th>
			<th>
				Wearable
			</th>			
		</tr>	
	</thead>
<% While (NOT rsGetMaterials.EOF) %>
    <tr>
      <td>
          <a class="btn btn-sm btn-danger mr-4 delete-material" data-tag-id="<%=(rsGetMaterials.Fields.Item("material_id").Value)%>"><i class="fa fa-trash-alt text-white"></i></a>
          <%=(rsGetMaterials.Fields.Item("material_name").Value)%>
	  </td>
      <td>
          <%If rsGetMaterials.Fields.Item("toggle_wearable").Value = true Then Response.Write "Yes" Else Response.Write "No"%>
	  </td>	  
    </tr>
<% 
rsGetMaterials.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsGetMaterials.Close()
Set rsGetMaterials = Nothing
%>
