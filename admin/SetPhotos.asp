<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../dwzPaging/dwzPaging.asp" -->
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
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
if Request.Querystring("Delete") = "Yes" then

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "DELETE FROM TBL_PhotoGallery WHERE PhotoID = " & Request.Querystring("PhotoID")
commUpdate.Execute()

end if
%>

<%
Dim rsSetPhotos__MMColParam
rsSetPhotos__MMColParam = "1"
If (Request("MM_EmptyValue") <> "") Then 
  rsSetPhotos__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsSetPhotos
Dim rsSetPhotos_cmd
Dim rsSetPhotos_numRows

Set rsSetPhotos_cmd = Server.CreateObject ("ADODB.Command")
rsSetPhotos_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsSetPhotos_cmd.CommandText = "SELECT PhotoID, ProductID, DetailID, filename, [description] FROM dbo.TBL_PhotoGallery WHERE status = ? AND DetailID = 0 ORDER BY PhotoID ASC" 
rsSetPhotos_cmd.Prepared = true
rsSetPhotos_cmd.Parameters.Append rsSetPhotos_cmd.CreateParameter("param1", 5, 1, -1, rsSetPhotos__MMColParam) ' adDouble

Set rsSetPhotos = rsSetPhotos_cmd.Execute
rsSetPhotos_numRows = 0
%>

<%
'*********************************
'*  RECORDSET PAGING - NUMERIC
'*  http://www.dwzone-it.com
'*  Version 1.1.0
'*********************************
set dwzPaging_0 = new dwzRecPaging
dwzPaging_0.init()
dwzPaging_0.setTypeNumeric()
dwzPaging_0.setRecordset rsSetPhotos
dwzPaging_0.setRecPaging 15
dwzPaging_0.setPages 5
dwzPaging_0.setStyle "HomePageLinks", "HomePageLinks", "HomePageLinks"
dwzPaging_0.setText "Next", "Previous", "First", "Last"
dwzPaging_0.setSeparator " | "
dwzPaging_0.setLinkMask "[ {1} ]"
dwzPaging_0.Execute()
set rsSetPhotos = dwzPaging_0.getRecordset()
%>
<%
Dim Repeat5__numRows
Dim Repeat5__index

Repeat5__numRows = 20
Repeat5__index = 0
rsSetPhotos_numRows = rsSetPhotos_numRows + Repeat5__numRows
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetProducts_numRows = rsGetProducts_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsSetPhotos_total
Dim rsSetPhotos_first
Dim rsSetPhotos_last

' set the record count
rsSetPhotos_total = rsSetPhotos.RecordCount

' set the number of rows displayed on this page
If (rsSetPhotos_numRows < 0) Then
  rsSetPhotos_numRows = rsSetPhotos_total
Elseif (rsSetPhotos_numRows = 0) Then
  rsSetPhotos_numRows = 1
End If

' set the first and last displayed record
rsSetPhotos_first = 1
rsSetPhotos_last  = rsSetPhotos_first + rsSetPhotos_numRows - 1

' if we have the correct record count, check the other stats
If (rsSetPhotos_total <> -1) Then
  If (rsSetPhotos_first > rsSetPhotos_total) Then
    rsSetPhotos_first = rsSetPhotos_total
  End If
  If (rsSetPhotos_last > rsSetPhotos_total) Then
    rsSetPhotos_last = rsSetPhotos_total
  End If
  If (rsSetPhotos_numRows > rsSetPhotos_total) Then
    rsSetPhotos_numRows = rsSetPhotos_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsSetPhotos_total = -1) Then

  ' count the total records by iterating through the recordset
  rsSetPhotos_total=0
  While (Not rsSetPhotos.EOF)
    rsSetPhotos_total = rsSetPhotos_total + 1
    rsSetPhotos.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsSetPhotos.CursorType > 0) Then
    rsSetPhotos.MoveFirst
  Else
    rsSetPhotos.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsSetPhotos_numRows < 0 Or rsSetPhotos_numRows > rsSetPhotos_total) Then
    rsSetPhotos_numRows = rsSetPhotos_total
  End If

  ' set the first and last displayed record
  rsSetPhotos_first = 1
  rsSetPhotos_last = rsSetPhotos_first + rsSetPhotos_numRows - 1
  
  If (rsSetPhotos_first > rsSetPhotos_total) Then
    rsSetPhotos_first = rsSetPhotos_total
  End If
  If (rsSetPhotos_last > rsSetPhotos_total) Then
    rsSetPhotos_last = rsSetPhotos_total
  End If

End If
%>
<%
' *** FX Update Multiple Records in FRM_UpdatePhotos
If (rsSetPhotos_first <> "") Then upd_rsSetPhotos = rsSetPhotos_first-1 Else upd_rsSetPhotos = 0 End If ' counter
If (cStr(Request.Form("SubmitActive")) <> "") Then
  FX_sqlerror = ""
  FX_updredir = "SetPhotos.asp"
  tmp = "ADODB.Command"
  Set update_Multi = Server.CreateObject(tmp)
  update_Multi.ActiveConnection = MM_bodyartforms_sql_STRING
  For N = upd_rsSetPhotos+1 To rsSetPhotos_total
      If (Request.Form("DetailID"&N) <> "") Then s1 = Replace(Request.Form("DetailID"&N),"'","''") Else s1 = "0" End If
      If (Request.Form("ProductID"&N) <> "") Then s2 = Replace(Request.Form("ProductID"&N),"'","''") Else s2 = "0" End If
      If (Request.Form("fx_updmatch"&N) <> "") Then sw = Replace(Request.Form("fx_updmatch"&N),"'","''") Else sw = "0" End If
    On Error Resume Next
      update_Multi.CommandText = "UPDATE dbo.TBL_PhotoGallery SET DetailID="+s1+", ProductID="+s2+" WHERE PhotoID="+sw+""
      update_Multi.Execute
    If (Err.Description <> "") Then
      FX_sqlerror = FX_sqlerror & "Row " & N & ": " & Err.Description & "<br><br>"
    End If
  Next
  update_Multi.ActiveConnection.Close
  thispath = cStr(Request.ServerVariables("SCRIPT_NAME"))
  If (FX_updredir = "") Then FX_updredir = Mid(thispath, InstrRev(thispath, "/")+1) End If
  If (Request.QueryString <> "") Then
    ch = "&"
    If (InStr(FX_updredir,"?") = 0) Then ch = "?" End If
    FX_updredir = FX_updredir & ch & Request.QueryString
  End If
  If (FX_sqlerror <> "") Then
    Response.Write("<font color=""red"">"&FX_sqlerror&"</font>")
  Else Response.Redirect(FX_updredir) End If
End If
%>
<html>
<head>

<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<link rel="stylesheet" type="text/css" href="../jquery.fancybox/jquery.fancybox.css" media="screen" />
	<script type="text/javascript" src="../jquery.fancybox/jquery-1.3.2.min.js"></script>
	<script type="text/javascript" src="../jquery.fancybox/jquery.easing.1.3.js"></script>
	<script type="text/javascript" src="../jquery.fancybox/jquery.fancybox-1.2.1.pack.js"></script>
<script type="text/javascript"> 
 
    $(document).ready(function(){
	
		$("a.fancy").fancybox();
		$("a.FancyFrame").fancybox({ 'frameWidth': 650, 'frameHeight': 450});
    }); 
 
</script>

<title>Set photo gauges</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body class="mainbkgd">
<!--#include file="admin_header.asp"-->
<P>
<form METHOD="POST" ACTION="<%=MM_editAction%>" name="FRM_UpdatePhotos" id="FRM_UpdatePhotos">    <table width="100%" border="0" cellpadding="5" cellspacing="1">
<tr class="checkoutHeader">
        <td colspan="2">Set photo details (<%=(rsSetPhotos_total)%> left)</td>
      </tr>
      <tr class="checkoutHeader">
        <td width="50%" align="left"><input name="SubmitActive" type="submit" id="SubmitActive" value="Update details" size="10" /></td>
        <td width="50%" align="right"><%
'*********************************
'*  RECORDSET PAGING - NUMERIC
'*********************************
dwzPaging_0.GetPaging()
'*********************************
'*  RECORDSET PAGING - NUMERIC
'*********************************
%></td>
      </tr>
      <% 
While ((Repeat5__numRows <> 0) AND (NOT rsSetPhotos.EOF)) 
%>
<% upd_rsSetPhotos = upd_rsSetPhotos+1 %>        <tr class="faqs">
          <td colspan="2" bgcolor="#ececec"><font color="#000000"><a href="http://www.bodyartforms.com/gallery/uploads/<%=(rsSetPhotos.Fields.Item("filename").Value)%>" rel="group" class="fancy"><img src="http://www.bodyartforms.com/gallery/uploads/thumb_<%=(rsSetPhotos.Fields.Item("filename").Value)%>" width="120" height="120" hspace="10" align="left" border="0" ></a><%=(rsSetPhotos.Fields.Item("description").Value)%>&nbsp;&nbsp;<br>
       
          <%
Dim rsGetProducts
Dim rsGetProducts_cmd
Dim rsGetProducts_numRows

Set rsGetProducts_cmd = Server.CreateObject ("ADODB.Command")
rsGetProducts_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetProducts_cmd.CommandText = "SELECT ProductDetailID, ProductID, ProductDetail1, Gauge, Length FROM dbo.ProductDetails WHERE ProductID = "&rsSetPhotos.Fields.Item("ProductID").Value&" ORDER BY item_order ASC, Price ASC" 
rsGetProducts_cmd.Prepared = true

Set rsGetProducts = rsGetProducts_cmd.Execute
rsGetProducts_numRows = 0
%>
 
      
              <select name="DetailID<%=upd_rsSetPhotos%>" style="font-size: 16px;">
                <% If Not rsGetProducts.EOF Or Not rsGetProducts.BOF Then %>
                <option value="0">None</option>  
                <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetProducts.EOF)) 
%>
                
                      <option value="<%=(rsGetProducts.Fields.Item("ProductDetailID").Value)%>"><%=(rsGetProducts.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetProducts.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetProducts.Fields.Item("ProductDetail1").Value)%></option>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetProducts.MoveNext()
Wend
%>
                <% End If ' end Not rsGetProducts.EOF Or NOT rsGetProducts.BOF %>
                <%
rsGetProducts.Close()
Set rsGetProducts = Nothing
%>          
              </select>
          <br />
       
          <input type="hidden" name="fx_updmatch<%=upd_rsSetPhotos%>" value="<%=(rsSetPhotos.Fields.Item("PhotoID").Value)%>">
          <input type="text" name="ProductID<%=upd_rsSetPhotos%>" value="<%=(rsSetPhotos.Fields.Item("ProductID").Value)%>" >
          <br />
              <a href="SetPhotos.asp?PhotoID=<%=(rsSetPhotos.Fields.Item("PhotoID").Value)%>&Delete=Yes">Delete ID <%=(rsSetPhotos.Fields.Item("PhotoID").Value)%></a>      
          </font></td>
        </tr>
        <% 
  Repeat5__index=Repeat5__index+1
  Repeat5__numRows=Repeat5__numRows-1
  rsSetPhotos.MoveNext()
Wend
%>    
      <tr class="materialText">
        <td colspan="2" align="center" valign="middle"></td>
      </tr>

    </table>
  <input name="SubmitActive" type="submit" id="SubmitActive" value="Update details" size="10" />
</form> 
</body>
</html>
<%
rsSetPhotos.Close()
Set rsSetPhotos = Nothing
%>

<%
'*********************************
'*  RECORDSET PAGING - NUMERIC
'*********************************
set dwzPaging_0 = nothing
'*********************************
'*  RECORDSET PAGING - NUMERIC
'*********************************
%>
