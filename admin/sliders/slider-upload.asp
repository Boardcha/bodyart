<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../Connections/chilkat.asp" -->
<!--#include file="../../Connections/aws-s3.asp" -->
<!--#include file="inc_content_type.asp" -->
<%

	Set Upload = Server.CreateObject("Persits.Upload")
	photo_path = "img_temp"
	Upload.OverwriteFiles = True
	Upload.SaveVirtual(photo_path)
	
	Set File = Upload.Files("file")

	If Not File Is Nothing Then
		imgDimension =  Upload.form("img_id")
		sliderId =  Upload.form("slider_id")
		fileName = File.fileName
		'File.Copy Server.MapPath(photo_path & "\" & fileName)
			
		set objFs=Server.CreateObject("Scripting.FileSystemObject")
		If objFs.FileExists(Server.MapPath(photo_path & "\" & fileName)) Then
			' Upload files to S3 bucket
			set http = Server.CreateObject("Chilkat_9_5_0.Http")
			' Insert your AWS keys here:
			http.AwsAccessKey = AWS_ACCESSKEY
			http.AwsSecretKey = AWS_SECRETKEY
			bucketName = "baf-hero-images"
			contentType = getContentTypeFromFileName(fileName) ' This is mandatory otherwise S3 cannot register mime type

			success = http.S3_UploadFile(Replace(Server.MapPath(photo_path & "\" & filename), "\", "/"), contentType, bucketName, fileName)

			If (success) Then
				Response.Write "{""isSuccessful"":""yes"",""imgName"":""" & fileName & """}"
			Else
				Response.Write "{""isSuccessful"":""no"",""error"":""" & Server.HTMLEncode( http.LastErrorText) & """}"
			End If
		End If
		File.delete	
	End If
	set objFs = nothing	
	Set Upload = nothing
	DataConn.Close()
%>

