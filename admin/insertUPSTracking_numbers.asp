<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!-- #include file="FXInc/clsNoComUpload.inc" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

' set object
Set upFile = New NoComUpload
%>
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Update Table from CSV with UploadTracking
If (cStr(upFile.Form("SUBMIT")) <> "" AND Trim(cStr(upFile.Form("FILE1"))) <> "") Then
  valid = false
  sql_errmsg = ""
  invalidmsg = "This is not a valid file"
  path = Server.MapPath("fx_temp")
  upFile.Save path,"FILE1",true
  fileurl = upFile.MakeFilePath(path, "FILE1")
  set fxfso = Server.CreateObject("Scripting.FileSystemObject")
  set fxfobj = fxfso.getfile(fileurl)
  filetype = fxfobj.type
  fileext = fxfso.GetExtensionName(fileurl)
  If (InStr(LCase(filetype),"csv") <> 0 OR InStr(LCase(filetype),"excel") <> 0 OR fileext = "txt") Then valid = true End If
  If (valid) Then
    fx_dbfields = array("UPS_tracking")
    fx_wherecol = "ID"
    tmp = "ADODB.Command"
    Set update_Csv = Server.CreateObject(tmp)
    update_Csv.ActiveConnection = MM_bodyartforms_sql_STRING
    set fxfile = fxfobj.OpenAsTextStream(1,-2)
    content = fxfile.ReadAll()
    lines = split(content, vbCrLf)
    For i=0 To UBound(lines)
      data = split(Replace(Replace(lines(i),"""",""),"'",""), ",")
      fxnum = -1
      cols = UBound(data)
      If (i = 0) Then csvCols = data End If
      If (i > 0 AND lines(i) <> "") Then
        If (cols >= UBound(fx_dbfields)) Then
          For df=0 To UBound(fx_dbfields)
            execute(fx_dbfields(df) & " = """"")
            For c=0 To cols
              If (LCase(csvCols(c)) = LCase(fx_dbfields(df))) Then
                execute(fx_dbfields(df) & " = data(c)")
                fxnum = fxnum+1
                Exit For
              End If
            Next
          Next
          For w=0 To cols
            If (LCase(csvCols(w)) = LCase(fx_wherecol)) Then
              execute(fx_wherecol & " = data(w)")
              fxnum = fxnum+1
              Exit For
            End If
          Next
          If (fxnum = UBound(fx_dbfields)+1) Then
            UPS_tracking = Replace(UPS_tracking,"'","''")
            ID = Replace(ID,"'","''")
            On Error Resume Next
              update_Csv.CommandText = "UPDATE dbo.sent_items SET UPS_tracking='"+UPS_tracking+"' WHERE ID="+ID+""
              update_Csv.Execute
            If (Err.Description <> "") Then
              sql_errmsg = sql_errmsg & "Row " & i & ": " & Err.Description & "<br><br>"
            End If
          Else
            csverrmsg = "The file doesn't contain all the expected columns"
            Exit For
          End If
        Else
          csverrmsg = invalidmsg
        End If
      End If
    Next
    fxfile.Close
    update_Csv.ActiveConnection.Close
  Else
    csverrmsg = invalidmsg
  End If
  If (fxfso.fileExists(fileurl)) Then fxfso.deleteFile fileurl End If
  set fxfso = nothing
  If (sql_errmsg <> "") Then csverrmsg = sql_errmsg End If
  If (csverrmsg = "") Then
    csvdone = true
  End If
End If
set upFile = Nothing
%>

<html>
<title>Insert UPS tracking #'s</title>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5> Insert UPS tracking #'s into database </h5>
CSV (UPS_Import_Numbers.csv):
<FORM class="form-inline" ACTION="<%=MM_editAction%>" METHOD="POST" enctype="multipart/form-data" name="UploadTracking" id="UploadTracking">
      <INPUT class="form-control form-control-sm" NAME="FILE1" TYPE="FILE" SIZE="40">
    <INPUT class="btn btn-purple ml-4" name="SUBMIT" TYPE=SUBMIT VALUE="Upload!">
  </p>
</FORM>
<% If (csvdone <> "") Then %>
<div class="alert alert-success"
   Tracking #'s imported successfully
</div> 
   <% End If ' errcsv %>
  
  <% If (csverrmsg <> "") Then Response.Write(csverrmsg) End If ' errcsv %>
</div>
</html>
