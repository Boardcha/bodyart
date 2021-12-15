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
	
	Set File_1 = Upload.Files("file[0]")
	Set File_2 = Upload.Files("file[1]")
	Set File_3 = Upload.Files("file[2]")
	
	If Not File_1 Is Nothing Then File1_ImgWidth = getImageWidth("img_temp\" & File_1.fileName)
	If Not File_2 Is Nothing Then File2_ImgWidth = getImageWidth("img_temp\" & File_2.fileName)	
	If Not File_3 Is Nothing Then File3_ImgWidth = getImageWidth("img_temp\" & File_3.fileName)	
	
	If File1_ImgWidth > 500 Then
		fileName = Left(File_1.FileName, InstrRev(File_1.FileName, ".") - 1) & "." & Mid(File_1.FileName, InstrRev(File_1.FileName, ".") + 1)
		thumbnailName = "thumbnail-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(File_1.FileName, InstrRev(File_1.FileName, ".") + 1)
	Elseif File2_ImgWidth > 500 Then	
		fileName = Left(File_2.FileName, InstrRev(File_2.FileName, ".") - 1) & "." & Mid(File_2.FileName, InstrRev(File_2.FileName, ".") + 1)
		thumbnailName = "thumbnail-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(File_2.FileName, InstrRev(File_2.FileName, ".") + 1)	
	Elseif File3_ImgWidth > 500 Then	
		fileName = Left(File_3.FileName, InstrRev(File_3.FileName, ".") - 1) & "." & Mid(File_3.FileName, InstrRev(File_3.FileName, ".") + 1)
		thumbnailName = "thumbnail-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(File_3.FileName, InstrRev(File_3.FileName, ".") + 1)	
	End If
		
	If Not File_1 Is Nothing Then
		if File1_ImgWidth < 100 Then '90x90
			File_1.Copy Server.MapPath(photo_path & "\90x90\" & thumbnailName)
		elseif File1_ImgWidth < 500 Then '400x400
			File_1.Copy Server.MapPath(photo_path & "\400x400\" & thumbnailName)
		elseif File1_ImgWidth > 500 Then '1000x1000
			File_1.Copy Server.MapPath(photo_path & "\1000x1000\" & fileName)
		end if
		File_1.Delete		
	End If
	
	If Not File_2 Is Nothing Then
		if File2_ImgWidth < 100 Then '90x90
			File_2.Copy Server.MapPath(photo_path & "\90x90\" & thumbnailName)
		elseif File2_ImgWidth < 500 Then '400x400
			File_2.Copy Server.MapPath(photo_path & "\400x400\" & thumbnailName)
		elseif File2_ImgWidth > 500 Then '1000x1000
			File_2.Copy Server.MapPath(photo_path & "\1000x1000\" & fileName)
		end if
		File_2.Delete		
	End If
	
	If Not File_3 Is Nothing Then
		if File3_ImgWidth < 100 Then '90x90
			File_3.Copy Server.MapPath(photo_path & "\90x90\" & thumbnailName)
		elseif File3_ImgWidth < 500 Then '400x400
			File_3.Copy Server.MapPath(photo_path & "\400x400\" & thumbnailName)
		elseif File3_ImgWidth > 500 Then '1000x1000
			File_3.Copy Server.MapPath(photo_path & "\1000x1000\" & fileName)
		end if
		File_3.Delete		
	End If	
	
	set objFs=Server.CreateObject("Scripting.FileSystemObject")
	If objFs.FileExists(Server.MapPath(photo_path & "\1000x1000\" & fileName)) AND objFs.FileExists(Server.MapPath(photo_path & "\400x400\" & thumbnailName)) AND objFs.FileExists(Server.MapPath(photo_path & "\90x90\" & thumbnailName)) Then
		' Upload files to S3 bucket
		set http = Server.CreateObject("Chilkat_9_5_0.Http")
		' Insert your AWS keys here:
		http.AwsAccessKey = AWS_ACCESSKEY
		http.AwsSecretKey = AWS_SECRETKEY
		bucketName = "bodyartforms-products"
		thumbnailBucketName = "baf-thumbs-400"
		contentType = getContentTypeFromFileName(fileName) ' This is mandatory otherwise S3 cannot register mime type

		success1 = http.S3_UploadFile(Replace(Server.MapPath(photo_path & "\1000x1000\" & filename), "\", "/"), contentType, bucketName, fileName)
		success2 = http.S3_UploadFile(Replace(Server.MapPath(photo_path & "\400x400\" & thumbnailName), "\", "/"), contentType, thumbnailBucketName, thumbnailName)
		success3 = http.S3_UploadFile(Replace(Server.MapPath(photo_path & "\90x90\" & thumbnailName), "\", "/"), contentType, bucketName, thumbnailName)

		If (success1 AND success2 AND success3) Then
			set objCmd = Server.CreateObject("ADODB.Command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE jewelry SET picture=?, largepic=? WHERE productID=?"
			objCmd.Parameters.Append(objCmd.CreateParameter("picture", 200, 1, 100, thumbnailName))
			objCmd.Parameters.Append(objCmd.CreateParameter("largepic", 200, 1, 100, fileName))
			objCmd.Parameters.Append(objCmd.CreateParameter("productID" ,3 ,1, 15, Upload.form("productid")))
			objCmd.Execute()
			Response.Write "https://s3.amazonaws.com/" & bucketName & "/" & thumbnailName
		Else
			Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"	
		End If
	End If
	
	' Delete images from temporary folder
	Set objFolder = objFS.GetFolder(Server.MapPath(photo_path)) 
	Set objFiles = objFolder.Files
	dim curFile
	For each curFile in objFiles
	'==== COMMENTED OUT SO THAT SERVER DOESNT DELETE SUBFOLDERS
	'	objFS.DeleteFile(curFile)
	Next	
	' Delete file in subfolder
	For Each subFolder In objFolder.SubFolders
		For each curFile in subFolder.Files
			objFS.DeleteFile(curFile)
		Next
	Next	
	
set objFs = nothing	
Set Upload = nothing
DataConn.Close()

Function getImageWidth(file)
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Path = Server.MapPath(file)
	Jpeg.Open Path	
	getImageWidth = Jpeg.OriginalWidth
	Set Jpeg = Nothing
End Function

%>