<%@ Language="VBSCRIPT" EnableSessionState=False %>
<%
Response.Expires = 60
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="incPU3Utils.asp" -->
<%
If Request.QueryString("UploadId") <> "" Then
  Dim progress, i
  Set progress = New UploadProgress
  progress.UploadId = Request.QueryString("UploadId")
  If Request.QueryString("Reset") = "true" Then
    progress.RemoveAll()
  End If
  Response.ContentType = "text/xml"
  Response.CharSet = "UTF-8"
  Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?>"
  Response.Write "<progressStatus totalBytes=""" & progress.TotalBytes & """"
  Response.Write " uploadedBytes=""" & progress.UploadedBytes & """"
  Response.Write " status=""" & progress.Status & """"
  Response.Write " lastFile=""" & progress.LastFile & """"
  Response.Write " lastError=""" & progress.LastError & """>"
  Response.Write "<files>"
  For i = 1 To progress.GetNumberOf("File")
    Response.Write "<file name=""" & progress.GetValue("File" & i) & """"
    Response.Write " status=""" & progress.GetValue("File" & i & "Status") & """"
    Response.Write " error=""" & progress.GetValue("File" & i & "Error") & """ />"
  Next
  Response.Write "</files>"
  Response.Write "<errors>"
  For i = 1 To progress.GetNumberOf("Error")
    Response.Write "<error description=""" & progress.GetValue("Error" & i) & """ />"
  Next
  Response.Write "</errors>"
  Response.Write "</progressStatus>"
Else
  ' No UploadId
End If
%>