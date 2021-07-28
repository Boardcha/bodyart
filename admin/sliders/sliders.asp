<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "FRM_DeleteSlider" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Sliders"
  MM_editColumn = "sliderID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "sliders.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl
    Else
      MM_editRedirectUrl = MM_editRedirectUrl
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
	' First, delete images inside the slider
  	Session("sliderId") = CStr(Request("MM_recordId"))
	Server.Execute "slider-delete-all-images.asp"
	
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	

	
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsGetSliders
Dim rsGetSliders_numRows

Set rsGetSliders = Server.CreateObject("ADODB.Recordset")
rsGetSliders.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetSliders.Source = "SELECT * FROM dbo.TBL_Sliders ORDER BY show_up_order ASC"
rsGetSliders.CursorLocation = 3 'adUseClient
rsGetSliders.LockType = 1 'Read-only records
rsGetSliders.Open()

rsGetSliders_numRows = 0
%>
<%
Dim rsEditSliders__MMColParam
rsEditSliders__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsEditSliders__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetSliders_numRows = rsGetSliders_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Manage homepage sliders</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- Daterangepicker -->
<link rel="stylesheet" href="../plugins/datepicker/daterangepicker.css" type="text/css">
<link rel="stylesheet" href="../css/dropzone.css" type="text/css">
<link rel="stylesheet" href="/CSS/jquery.fancybox.min.css" />
</head>
<body>
<!--#include file="../admin_header.asp"-->
<div class="pt-3" style="max-width:950px">

<% if request.querystring("Add") = "yes" then %>
<div class="pl-3">
<h5>Add a new slider</h5>
<div id="sliders">
	<div class="form-group">
		<div class="form-row">
			<div class="col-sm mt-3">
				<form action="#" class="dropzone single mr-3 img-upload" id="dropzone_1" data-img-id="550x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">550 x 350<span class="failed"></span></span></span></div>
				</form>
				<div class="text-center pr-2">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-img-id="550x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>
			<div class="col-sm mt-3">			
				<form action="#" class="dropzone single mr-3 img-upload" id="dropzone_2" data-img-id="850x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">850 x 350<span class="failed"></span></span></span></div>
				</form>
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-img-id="850x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>	
			<div class="col-sm mt-3">			
				<form action="#" class="dropzone single mr-3 img-upload" id="dropzone_3" data-img-id="1024x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">1024 x 350<span class="failed"></span></span></span></div>
				</form>
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-img-id="1024x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>	
			<div class="col-sm mt-3">			
				<form action="#" class="dropzone single mr-3 img-upload" id="dropzone_4" data-img-id="1600x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">1600 x 350<span class="failed"></span></span></span></div>
				</form>
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-img-id="1600x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>	
			<div class="col-sm mt-3">
				<form action="#" class="dropzone single mr-3 img-upload" id="dropzone_5" data-img-id="1920x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">1920 x 350<span class="failed"></span></span></span></div>
				</form>	
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-img-id="1920x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>
		</div>
	</div>
</div>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_AddSlider" id="FRM_AddSlider">
	<table class="table table-sm table-borderless w-100">
		<tr>
			<td width="20%" class="align-middle">Title (Internal use only)&nbsp; </td>
			<td width="80%"><input name="friendly_title" type="text" class="form-control form-control-sm w-75" id="friendly_title"></td>
		  </tr>
		<tr>
		  <td width="20%" class="align-middle">HTML Alt tag name&nbsp; </td>
		  <td width="80%"><input name="slider_name" type="text" class="form-control form-control-sm w-75" id="slider_name"></td>
		</tr>
	
		<tr>
		  <td width="20%" class="align-middle">URL</td>
		  <td width="80%"><input name="url" type="text" class="form-control form-control-sm w-75" id="url"></td>
		</tr>		
		<tr>
		  <td width="20%" class="align-middle">Date Start</td>
		  <td width="80%">		
			<input name="date_start" id="date_start" type="text" class="form-control form-control-sm w-25" value="" autocomplete="off" />
		  </td>
		</tr>
		<tr>
		  <td width="20%" class="align-middle">Date End</td>
		  <td width="80%">
			<input name="date_end" id="date_end" type="text" class="form-control form-control-sm w-25" value="" autocomplete="off" />
		  </td>
		</tr>		
		<tr>
		  <td width="20%">&nbsp;</td>
		  <td width="80%"><input class="btn btn-secondary" type="submit" name="Submit" value="Submit"></td>
		</tr>
	  </table>	

<input type="hidden" name="MM_insert" value="FRM_AddSlider">
<input type="hidden" name="img550x350" value="">
<input type="hidden" name="img850x350" value="">
<input type="hidden" name="img1024x350" value="">
<input type="hidden" name="img1600x350" value="">
<input type="hidden" name="img1920x350" value="">
</form>
</div>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "FRM_AddSlider") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Sliders"
  MM_editRedirectUrl = "sliders.asp"
  MM_fieldsStr  = "slider_name|value|friendly_title|value|date_start|value|date_end|value|url|value|img550x350|value|img850x350|value|img1024x350|value|img1600x350|value|img1920x350|value"
  MM_columnsStr = "slider_name|',none,''|friendly_title|',none,''|date_start|',none,''|date_end|',none,''|url|',none,'',''|img550x350|',none,'',''|img850x350|',none,'',''|img1024x350|',none,'',''|img1600x350|',none,'',''|img1920x350|',none,''"  

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl
    Else
      MM_editRedirectUrl = MM_editRedirectUrl
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "INSERT INTO " & MM_editTable & " (" & MM_tableValues & ", show_up_order) SELECT " & MM_dbValues & ", COALESCE(MAX(show_up_order),0)+1 FROM " & MM_editTable

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<% end if %>
<% if request.querystring("Edit") = "yes" then %>
<%
Dim rsEditSliders
Dim rsEditSliders_numRows

Set rsEditSliders = Server.CreateObject("ADODB.Recordset")
rsEditSliders.ActiveConnection = MM_bodyartforms_sql_STRING
rsEditSliders.Source = "SELECT * FROM dbo.TBL_Sliders WHERE sliderID = " + Replace(rsEditSliders__MMColParam, "'", "''") + ""
rsEditSliders.CursorLocation = 3 'adUseClient
rsEditSliders.LockType = 1 'Read-only records
rsEditSliders.Open()

rsEditSliders_numRows = 0
If Not IsNull(rsEditSliders.Fields.Item("img550x350").Value) AND rsEditSliders.Fields.Item("img550x350").Value <> "" Then
	img550x350_exist = true
End If
If Not IsNull(rsEditSliders.Fields.Item("img850x350").Value) AND rsEditSliders.Fields.Item("img850x350").Value <> "" Then
	img850x350_exist = true
End If
If Not IsNull(rsEditSliders.Fields.Item("img1024x350").Value) AND rsEditSliders.Fields.Item("img1024x350").Value <> "" Then
	img1024x350_exist = true
End If
If Not IsNull(rsEditSliders.Fields.Item("img1600x350").Value) AND rsEditSliders.Fields.Item("img1600x350").Value <> "" Then
	img1600x350_exist = true
End If
If Not IsNull(rsEditSliders.Fields.Item("img1920x350").Value) AND rsEditSliders.Fields.Item("img1920x350").Value <> "" Then
	img1920x350_exist = true
End If
%>
<div class="pl-3">
<h5>Edit Slider</h5>
<i>Click any image to enlarge</i>
<div id="sliders">
	<div class="form-group">
		<div class="form-row">
			<div class="col-sm mt-3">
				<form action="#" class="dropzone single mr-3 img-upload <%If img550x350_exist Then Response.Write "img-loaded"%>" id="dropzone_1" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="550x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">550 x 350<span class="failed"></span></span></span></div>
					<%If img550x350_exist Then%>
					<div class="dz-preview dz-processing dz-image-preview dz-success dz-complete">
						<div class="dz-image">
							<a style="pointer-events:initial" class="position-relative pointer" data-fancybox="product-images" data-caption="550 x 350" href="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img550x350").Value%>" id="img_550x350">
								<img src="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img550x350").Value%>" />
							</a>
						</div>
					</div>
					<%End If%>					
				</form>
				<div class="text-center pr-2">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="550x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>
			<div class="col-sm mt-3">			
				<form action="#" class="dropzone single mr-3 img-upload <%If img850x350_exist Then Response.Write "img-loaded"%>" id="dropzone_2" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="850x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">850 x 350<span class="failed"></span></span></span></div>
					<%If img850x350_exist Then%>
					<div class="dz-preview dz-processing dz-image-preview dz-success dz-complete">
						<div class="dz-image">
							<a style="pointer-events:initial" class="position-relative pointer" data-fancybox="product-images" data-caption="850 x 350" href="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img850x350").Value%>" id="img_850x350">
								<img src="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img850x350").Value%>" />
							</a>
						</div>
					</div>
					<%End If%>
				</form>
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="850x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>	
			<div class="col-sm mt-3">			
				<form action="#" class="dropzone single mr-3 img-upload <%If img1024x350_exist Then Response.Write "img-loaded"%>" id="dropzone_3" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="1024x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">1024 x 350<span class="failed"></span></span></span></div>
					<%If img1024x350_exist Then%>
					<div class="dz-preview dz-processing dz-image-preview dz-success dz-complete">
						<div class="dz-image">
							<a style="pointer-events:initial" class="position-relative pointer" data-fancybox="product-images" data-caption="1024 x 350" href="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img1024x350").Value%>" id="img_1024x350">
								<img src="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img1024x350").Value%>" />
							</a>
						</div>
					</div>
					<%End If%>					
				</form>
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="1024x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>	
			<div class="col-sm mt-3">			
				<form action="#" class="dropzone single mr-3 img-upload <%If img1600x350_exist Then Response.Write "img-loaded"%>" id="dropzone_4" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="1600x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">1600 x 350<span class="failed"></span></span></span></div>
					<%If img1600x350_exist Then%>
					<div class="dz-preview dz-processing dz-image-preview dz-success dz-complete">
						<div class="dz-image">
							<a style="pointer-events:initial" class="position-relative pointer" data-fancybox="product-images" data-caption="1600 x 350" href="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img1600x350").Value%>" id="img_1600x350">
								<img src="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img1600x350").Value%>" />
							</a>
						</div>
					</div>
					<%End If%>					
				</form>
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="1600x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>	
			<div class="col-sm mt-3">
				<form action="#" class="dropzone single mr-3 img-upload <%If img1920x350_exist Then Response.Write "img-loaded"%>" id="dropzone_5" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="1920x350">
					<div class="dz-message" data-dz-message><span><i class="fa fa-image"></i><br><span class="img-label">1920 x 350<span class="failed"></span></span></span></div>
					<%If img1920x350_exist Then%>
					<div class="dz-preview dz-processing dz-image-preview dz-success dz-complete">
						<div class="dz-image">
							<a style="pointer-events:initial" class="position-relative pointer" data-fancybox="product-images" data-caption="1920 x 350" href="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img1920x350").Value%>" id="img_1920x350">
								<img src="https://sliders.bodyartforms.com/<%=rsEditSliders.Fields.Item("img1920x350").Value%>" />
							</a>
						</div>
					</div>
					<%End If%>					
				</form>	
				<div class="text-center">
					<button class="btn btn-sm btn-danger clear-dropzone" type="button" data-slider-id="<%=rsEditSliders.Fields.Item("sliderID").Value%>" data-img-id="1920x350"> <i class="fa fa-trash-alt" aria-hidden="true"></i> </button>	
				</div>	
			</div>
		</div>
	</div>
</div>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_EditSlider" id="FRM_EditSlider">
	<table class="table table-sm table-borderless w-100">
		<tr>
			<td width="20%" class="align-middle">Title (Internal use only)&nbsp; </td>
			<td width="80%"><input name="friendly_title" type="text" class="form-control form-control-sm w-75" id="friendly_title" value="<%=rsEditSliders.Fields.Item("friendly_title").Value%>"></td>
		  </tr>
		<tr>
		  <td width="20%" class="align-middle">HTML Alt tag name&nbsp; </td>
		  <td width="80%"><input name="slider_name" type="text" class="form-control form-control-sm w-75" id="slider_name" value="<%=rsEditSliders.Fields.Item("slider_name").Value%>"></td>
		</tr>
	
		<tr>
		  <td width="20%" class="align-middle">URL</td>
		  <td width="80%"><input name="url" type="text" class="form-control form-control-sm w-75" id="url" value="<%=rsEditSliders.Fields.Item("url").Value%>"></td>
		</tr>		
		<tr>
		  <td width="20%" class="align-middle">Date Start</td>
		  <td width="80%">		
			<input name="date_start" id="date_start" type="text" class="form-control form-control-sm w-25" value="<%=FormatDateTime(rsEditSliders.Fields.Item("date_start").Value, 2)%>" autocomplete="off" />
		  </td>
		</tr>
		<tr>
		  <td width="20%" class="align-middle">Date End</td>
		  <td width="80%">
			<input name="date_end" id="date_end" type="text" class="form-control form-control-sm w-25" value="<%=FormatDateTime(rsEditSliders.Fields.Item("date_end").Value, 2)%>" autocomplete="off" />
		  </td>
		</tr>		
		<tr>
		  <td width="20%">&nbsp;</td>
		  <td width="80%"><input class="btn btn-secondary" type="submit" name="Submit" value="Submit"></td>
		</tr>
	  </table>	  

<input type="hidden" name="MM_update" value="FRM_EditSlider">
<input type="hidden" name="MM_recordId" value="<%= rsEditSliders.Fields.Item("sliderID").Value %>">
<input type="hidden" name="img550x350" value="<%= rsEditSliders.Fields.Item("img550x350").Value %>">
<input type="hidden" name="img850x350" value="<%= rsEditSliders.Fields.Item("img850x350").Value %>">
<input type="hidden" name="img1024x350" value="<%= rsEditSliders.Fields.Item("img1024x350").Value %>">
<input type="hidden" name="img1600x350" value="<%= rsEditSliders.Fields.Item("img1600x350").Value %>">
<input type="hidden" name="img1920x350" value="<%= rsEditSliders.Fields.Item("img1920x350").Value %>">
</form>

</div>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "FRM_EditSlider" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Sliders"
  MM_editColumn = "sliderID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "sliders.asp"
  MM_fieldsStr  = "slider_name|value|friendly_title|value|date_start|value|date_end|value|url|value|img550x350|value|img850x350|value|img1024x350|value|img1600x350|value|img1920x350|value"
  MM_columnsStr = "slider_name|',none,''|friendly_title|',none,''|date_start|',none,''|date_end|',none,''|url|',none,'',''|img550x350|',none,'',''|img850x350|',none,'',''|img1024x350|',none,'',''|img1600x350|',none,'',''|img1920x350|',none,''"  

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl
    Else
      MM_editRedirectUrl = MM_editRedirectUrl
    End If
  End If

End If
%>

<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

	'If not all size of images are provided, de-active the slider
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_Sliders SET active=0 WHERE sliderID = ? AND (img550x350='' OR img850x350='' OR img1024x350='' OR img1600x350='' OR img1920x350='' OR img550x350 IS NULL OR img850x350 IS NULL OR img1024x350 IS NULL OR img1600x350 IS NULL OR img1920x350 IS NULL)" 
	objCmd.Parameters.Append(objCmd.CreateParameter("sliderID",3,1,10,CStr(Request("MM_recordId"))))
	Set rs_getImage_Filename = objCmd.Execute()

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<% end if %>
<% if request.querystring("Delete") = "yes" then %>
<div class="pt-3 pl-3">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_DeleteSlider" id="FRM_DeleteSlider">
  <div class="alert alert-danger">
  Confirm deletion: <b><i><%= request.querystring("name") %></i>. (All images in this slider will be deleted from the server.)</b><br/>
  
  <input class="btn btn-sm btn-danger mt-2" type="submit" name="Submit2" value="DELETE">
</div>

    <input type="hidden" name="MM_delete" value="FRM_DeleteSlider">
    <input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">

</form>
</div>
<% end if %>
<div class="pt-3">
	<span class="d-inline-block w-25">
		<a class="btn btn-primary ml-3 mb-2" href="sliders.asp?Add=yes" class="HomePageLinks"><i class="fa fa-plus-circle mr-2"></i>Add New</a>
	</span>
	<span class="d-inline-block w-75 text-right float-right">
		<i>(drag-and drop rows to change display order of sliders.)</i>
	</span>
<table id="slider-list" class="table table-striped table-borderless table-hover w-100">
<thead class="thead-dark">
<tr>
	<th scope="col">Sliders</th>
	<th scope="col" class="pl-3">&nbsp;Active</th>
	<th scope="col">Date Start</th>
	<th scope="col">Date End</th>
</tr>
</thead>
<tbody>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetSliders.EOF))
allImagesProvided = true
If IsNull(rsGetSliders.Fields.Item("img550x350").Value) OR rsGetSliders.Fields.Item("img550x350").Value="" Then allImagesProvided = false
If IsNull(rsGetSliders.Fields.Item("img850x350").Value) OR rsGetSliders.Fields.Item("img850x350").Value="" Then allImagesProvided = false
If IsNull(rsGetSliders.Fields.Item("img1024x350").Value) OR rsGetSliders.Fields.Item("img1024x350").Value="" Then allImagesProvided = false
If IsNull(rsGetSliders.Fields.Item("img1600x350").Value) OR rsGetSliders.Fields.Item("img1600x350").Value="" Then allImagesProvided = false
If IsNull(rsGetSliders.Fields.Item("img1920x350").Value) OR rsGetSliders.Fields.Item("img1920x350").Value="" Then allImagesProvided = false

%>
	<%If rsGetSliders("active") Then checked = "checked" Else checked=""%>
    <tr id="<%=(rsGetSliders.Fields.Item("sliderID").Value)%>" data-id="<%=(rsGetSliders.Fields.Item("sliderID").Value)%>">
      <td>
          <a class="btn btn-sm btn-danger mr-4" href="sliders.asp?Delete=yes&ID=<%=rsGetSliders.Fields.Item("sliderID").Value%>&name=<%=rsGetSliders.Fields.Item("friendly_title").Value%>"><i class="fa fa-trash-alt"></i></a>

          <a href="sliders.asp?Edit=yes&ID=<%=(rsGetSliders.Fields.Item("sliderID").Value)%>" target="_top" class="LeftNavLinks">
			<img style="border-radius:6px;object-fit: cover;width:100px;height:100px" class="mr-2" src="https://sliders.bodyartforms.com/<%=rsGetSliders.Fields.Item("img550x350").Value%>" />  
			<%=(rsGetSliders.Fields.Item("friendly_title").Value)%></a>
		</td>
      <td>
		<div class="onoffswitch small-green" <%If Not allImagesProvided Then Response.Write " title=""It cannot be activated until all images are provided for this slider."" style=""opacity:0.3"""%>">
			<input type="checkbox" class="onoffswitch-checkbox activate" id="activation-<%=rsGetSliders.Fields.Item("sliderID").Value%>" <%If Not allImagesProvided Then Response.Write "data-updatable=""false"" disabled "%> data-id="<%=rsGetSliders.Fields.Item("sliderID").Value%>" <%=checked%>>
			<label class="onoffswitch-label" style="margin-bottom:0;margin-top:0" for="activation-<%=rsGetSliders.Fields.Item("sliderID").Value%>">
				<span class="onoffswitch-inner"></span>
				<span class="onoffswitch-switch"></span>
			</label>
		</div>	  
	  </td>
	  <td><%=FormatDateTime(rsGetSliders.Fields.Item("date_start").Value, 2)%></td>
	  <td><%=FormatDateTime(rsGetSliders.Fields.Item("date_end").Value, 2)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetSliders.MoveNext()
Wend
%>
</tbody>
</table>

</div>
</div>
<style>
/* -------------- START SLIDER PAGE STYLES --------------------- */
.tDnD_whileDrag td {
    background-color: #ceecf5 !important;
    border-bottom: 2px solid #37a5c5;
}
.table-hover tbody tr:nth-child(odd):hover td{
	background-color: #f2f2f2;
}
.table-hover tbody tr:nth-child(even):hover td{
	background-color: #fff;
}
.onoffswitch.small-green {
    width: 74px;
	height: 28px;
}	
.onoffswitch.small-green .onoffswitch-inner:before {
	content: "YES";
	background-color: #75b936;
	color: #fff;
	padding-right: 12px;
	height: 28px;
	line-height: 28px;
}
.onoffswitch.small-green .onoffswitch-inner:after {
	content: "NO";
	background-color: #b2b2b2;
	color: #fff;
	padding-right: 12px;
	height: 28px;
	line-height: 28px;
}
.onoffswitch.small-green .onoffswitch-label{
	border: none;
    border-radius: 50px;
}	
.onoffswitch.small-green .onoffswitch-switch{
	width: 17px;
	right: 45px;
	box-shadow: 0 1px 2px 0 rgb(0 0 0 / 20%), 0 3px 4px 0 rgb(0 0 0 / 10%);
	margin: 3px;
    width: 22px;
    height: 22px;
    border-radius: 100%;
	border: none;
    background: #fff;	
}
.img-loaded{
	border-color: #59e859 !important;
	border-style: solid;
	pointer-events: none;
	cursor: default;
}
.img-loaded i{
	display:none
}
.clear-dropzone{
	display: none; 
	margin: auto;
}
.img-loaded ~ Div .clear-dropzone{
	display: block
}
.img-label{
	background-color: #fcfcfc; 
	padding: 2px 5px; 
	opacity: 0.8; 
	border-radius: 4px; 
	font-weight:600;
}
#sliders .dropzone{
	padding:0; 
	width:157px; 
	height: 157px; 
	border-color: #d3d3d3; 
	background-color: #fcfcfc;
}
#sliders .dropzone span{
	color:#696986
}
#sliders .dropzone i{
	color: #696986; 
	font-size: 25px;
}
#sliders .dropzone .dz-message {
    margin-top: 50px; 
    position: relative;
}
.img-loaded .dz-message {
    position: relative;
	z-index: 12;
}	
#sliders .dropzone .dz-preview{
	margin: 0; 
	position: absolute; 
	top: 2px;
}
#sliders .dropzone .dz-preview .dz-image {
    border-radius: 2px;
    overflow: hidden;
    width: 153px;
    height: 153px;
}	

#sliders .dropzone .dz-preview .dz-image img{
	object-fit: cover;
}

#sliders .dropzone .img-label.failed{
	content: '<br><span style="color:red">(Failed)</span>';
}

#sliders .dropzone .dz-preview:hover .dz-image img {
    -webkit-transform: none;
    -moz-transform: none;
    -ms-transform: none;
    -o-transform: none;
    transform: none;
    -webkit-filter: none;
    filter: none;
}
/* -------------- END SLIDER PAGE STYLES --------------------- */
</style>
<!-- Daterangepicker -->
<script src="../plugins/datepicker/daterangepicker.js"></script>
<script src="../plugins/tablednd/jquery.tablednd.js"></script>
<script src="../scripts/dropzone.js"></script>
<script src="/js/jquery.fancybox.min.js"></script>
<script>
Dropzone.autoDiscover = false;
$(document).ready(function () {
	// Initialize date pickers
	$('#date_start, #date_end').daterangepicker({
		singleDatePicker: true,
		autoUpdateInput: false,
		locale: {format: 'M/D/YYYY'}		
	});
	$('#date_start, #date_end').on('apply.daterangepicker', function (ev, picker) {
		$(this).val(picker.startDate.format('M/D/YYYY'));
	});
	$('#date_start, #date_end').on('cancel.daterangepicker', function (ev, picker) {
		$(this).val('');
	});
	
	// tableDnd plugin drag & drop table rows 
    $("#slider-list").tableDnD({sensitivity:12, onDrop: function(table, row) {
        var serial = $.tableDnD.serialize();
		var IDs = serial.split("&");
		var arrData = [];
		for(var key in IDs)
		{
			arrData.push(IDs[key].split("=")[1]);
		}

		$.ajax({
			method: "POST",
			url: "slider-reorder.asp",
			data: {data: arrData.join(',')}
		})		
    }});

	// Activate slider
	$(".activate").on("click", function () {
		var checked;
		var data_id = $(this).attr("data-id");
		var updatable = $(this).attr("data-updatable");
		if(updatable != "false"){
			if ($(this).is(":checked")) {
				checked = 1;
			} else {
				checked = 0;
			}
			$.ajax({
				method: "POST",
				url: "slider-activate.asp",
				data: {active: checked, id: data_id}
			})		
		}
	});
	
	// Dropzone - upload handling
	var dropzone = $("form.dropzone").dropzone({
		autoProcessQueue: true,
		parallelUploads : 1,
		uploadMultiple: false,
		url: "slider-upload.asp",
		init: function () {
			var myDropzone = this;
					
			myDropzone.on("thumbnail", function(file) {
				var arr = myDropzone.element.dataset.imgId.split("x");
				var target_width = parseInt(arr[0]);
				var target_height = parseInt(arr[1]);
				if (file.accepted !== false) {
					if (file.width != target_width || file.height != target_height) {
						myDropzone.removeFile(file);
						alert("Acceptable image dimensions for this area is " + target_width + " x " + target_height + ". Yours is " + file.width + " x " + file.height + ".");
					}
				}
            });
			
			myDropzone.on('sending', function(file, xhr, formData) {
				var sliderId = $("#" + myDropzone.element.id).attr("data-slider-id");
				var imgId = $("#" + myDropzone.element.id).attr("data-img-id");
				formData.append('slider_id', sliderId);
				formData.append('img_id', imgId);
			});
			
			myDropzone.on("error", function(file, response) {
				var imgId = $("#" + myDropzone.element.id).attr("data-img-id");			
				$("#" + myDropzone.element.id).css({"border-color":"red","border-style":"solid"});
				myDropzone.removeFile(file);
				$("#" + myDropzone.element.id + " .img-label .failed").html('<br><span style="color:red">(Failed)</span>');
				$("input[name='img" + imgId + "']").val('');
			});

			myDropzone.on('success', function(file, response) {
				var json = jQuery.parseJSON(response);
				var imgId = $("#" + myDropzone.element.id).attr("data-img-id");
				if(json.isSuccessful == "yes"){
					$("#" + myDropzone.element.id).removeClass("dz-started");
					$("#" + myDropzone.element.id).addClass("img-loaded");
					$("#" + myDropzone.element.id + " .img-label .failed").html('');
					$("input[name='img" + imgId + "']").val(json.imgName);
					$("#" + myDropzone.element.id + " .dz-image").html(
						'<a style="pointer-events:initial" class="position-relative pointer" data-fancybox="product-images" data-caption="' + imgId + '" href="https://sliders.bodyartforms.com/' + json.imgName + '" id="img_' + imgId + '">' +
							'<img src="https://sliders.bodyartforms.com/' + json.imgName + '" />' +
						'</a>'
					);		
				}else{
					$("#" + myDropzone.element.id).css({"border-color":"red","border-style":"solid"});
					myDropzone.removeFile(file);
					$("#" + myDropzone.element.id + " .img-label .failed").html('<br><span style="color:red">(Failed)</span>');
					$("input[name='img" + imgId + "']").val('');
				}
			});					
		}		
	});	
	
	// Delete image
	$(".clear-dropzone").click(function (e) {
		var sliderId = $(this).attr("data-slider-id");
		var imgId = $(this).attr("data-img-id");
		$.ajax({
			method: "POST",
			url: "slider-delete-image.asp",
			data: {slider_id: sliderId, img_id: imgId}
		});
		$("input[name='img" + imgId + "']").val('');
		dropzone_id = $('.dropzone[data-img-id=' + imgId + ']')[0].id;
		$("#" + dropzone_id + " .dz-image").remove();
		$("#" + dropzone_id).removeClass("img-loaded");
	});	
});
</script>	

</body>
</html>
<%
rsGetSliders.Close()
Set rsGetSliders = Nothing
%>
