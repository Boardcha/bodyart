<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
response.write "id: " & request.form("PhotoID")
filename = request.form("photo_filename")
  
if request.form("photo_status") = "1" then


	' Update photos status
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_PhotoGallery SET status = 1, ProductID = ?, reviewed_date = ?, reviewed_by = ? WHERE PhotoID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("@ProductID",3,1,10,request.form("ProductID")))
	objCmd.Parameters.Append(objCmd.CreateParameter("reviewed_date",200,1,20, FormatDateTime(date(),2) ))
	objCmd.Parameters.Append(objCmd.CreateParameter("reviewed_by",200,1,40, user_name ))
	objCmd.Parameters.Append(objCmd.CreateParameter("@PhotoID",3,1,10,request.form("PhotoID")))
	objCmd.Execute()

	' add points to customer account
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET Points = Points + 3 WHERE customer_ID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("@CustomerID",3,1,10,request.form("customerID")))
	objCmd.Execute()

	' Delete original temp file (full uncompressed photo) from \gallery\temp-uploads\
	DeleteFile = "\gallery\temp-uploads\" & filename
	DeleteFilePath = Server.MapPath(DeleteFile)
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(DeleteFilePath) then
		fs.DeleteFile(DeleteFilePath)
	end if
	set fs=nothing



else ' photo was rejected

  
	' Delete photo from database
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_PhotoGallery WHERE PhotoID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("@PhotoID",3,1,10,request.form("PhotoID")))
	objCmd.Execute()

	
	
	' Delete original temp file (full uncompressed photo) from \gallery\temp-uploads\
'	DeleteTemp = "\gallery\temp-uploads\" & filename
'	DeleteTempFile = Server.MapPath(DeleteTemp)
'	DeleteResizedFile = "\gallery\uploads\" & filename
'	DeleteResizedPath = Server.MapPath(DeleteResizedFile)
'	DeleteThumb = "\gallery\uploads\thumb_" & filename
'	DeleteThumbPath = Server.MapPath(DeleteThumb)
'	response.write "<br/>" & DeleteTempFile
'	response.write "<br/>" & DeleteResizedPath
'	response.write "<br/>" & DeleteThumbPath
'	Set fs=Server.CreateObject("Scripting.FileSystemObject")
'	if fs.FileExists(DeleteTempFile) then
'		fs.DeleteFile DeleteTempFile, true
'	end if
'	if fs.FileExists(DeleteResizedPath) then
'		fs.DeleteFile DeleteResizedPath, true
'	end if
'	if fs.FileExists(DeleteThumbPath) then
'		fs.DeleteFile DeleteThumbPath, true
'	end if

'	set fs=nothing	
	
	
if request.form("photo_status") <> "0" then ' send the email if it's got a reason selected

mailer_type = "reject-photo"
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->
<%
end if ' send the email if it's got a rejected reason


end if %>	

