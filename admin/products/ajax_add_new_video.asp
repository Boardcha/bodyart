<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../Connections/chilkat.asp" -->
<!--#include file="../../Connections/aws-s3.asp" -->
<!--#include file="inc_content_type.asp" -->
<%
	Dim video, img90, img400
	Set Upload = Server.CreateObject("Persits.Upload")
	'Upload.IgnoreNoPost = True
	photo_path = "img_temp"
	Upload.OverwriteFiles = True
	Upload.SaveVirtual(photo_path)
	
	arrFiles = Array(Upload.Files("file[0]"), Upload.Files("file[1]"), Upload.Files("file[2]"))
	
	For each file in arrFiles
		If Split(getContentTypeFromfileName(file.fileName), "/")(0) = "image" Then 
			If getImageWidth("img_temp\" & file.fileName) < 100 Then Set img90 = file
			If getImageWidth("img_temp\" & file.fileName) > 100 Then Set img400 = file	
		ElseIf Split(getContentTypeFromfileName(file.fileName), "/")(0) = "video" Then 
			Set video = file
		End If	
	Next
	
	If Not img90 Is Nothing  AND Not img400 Is Nothing  AND Not video Is Nothing Then

		thumbnailName = "thumbnail-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(img400.FileName, InstrRev(img400.FileName, ".") + 1)
		videoName = "video-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(video.FileName, InstrRev(video.FileName, ".") + 1)
		'Rename Files
		img90.Copy Server.MapPath(photo_path & "\90x90\" & thumbnailName)
		img90.Delete
		img400.Copy Server.MapPath(photo_path & "\400x400\" & thumbnailName)
		img400.Delete
		video.Copy Server.MapPath(photo_path & "\" & videoName)
		video.Delete		
	
		set objFs=Server.CreateObject("Scripting.FileSystemObject")
		If objFs.FileExists(Server.MapPath(photo_path & "\90x90\" & thumbnailName)) AND objFs.FileExists(Server.MapPath(photo_path & "\400x400\" & thumbnailName)) AND objFs.FileExists(Server.MapPath(photo_path & "\" & videoName)) Then
			' Upload files to S3 bucket
			set http = Server.CreateObject("Chilkat_9_5_0.Http")
			' Insert your AWS keys here:
			http.AwsAccessKey = AWS_ACCESSKEY
			http.AwsSecretKey = AWS_SECRETKEY
			thumnail_90x90_BucketName = "bodyartforms-products"
			thumnail_400x400_BucketName = "baf-thumbs-400"
			videoBucketName = "baf-videos"

			success1 = http.S3_UploadFile(Replace(Server.MapPath(photo_path & "\" & videoName), "\", "/"), getContentTypeFromfileName(videoName), videoBucketName, videoName)
			success2 = http.S3_UploadFile(Replace(Server.MapPath(photo_path & "\90x90\" & thumbnailName), "\", "/"), getContentTypeFromfileName(thumbnailName), thumnail_90x90_BucketName, thumbnailName)
			success3 = http.S3_UploadFile(Replace(Server.MapPath(photo_path & "\400x400\" & thumbnailName), "\", "/"), getContentTypeFromfileName(thumbnailName), thumnail_400x400_BucketName, thumbnailName)

			If (success1 AND success2 AND success3) Then
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "INSERT INTO tbl_images(product_id, img_full, img_thumb, is_video) VALUES (?, ?, ?, 1)"
				objCmd.Parameters.Append(objCmd.CreateParameter("productid" ,3 ,1, 15, Upload.form("productid")))
				objCmd.Parameters.Append(objCmd.CreateParameter("img_full", 200, 1, 100, videoName))
				objCmd.Parameters.Append(objCmd.CreateParameter("img_thumb", 200, 1, 100, thumbnailName))
				objCmd.Execute()
			Else
				Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"	
			End If
		End If
	End If
	
	' Delete images from temporary folder
	Set objFolder = objFS.GetFolder(Server.MapPath(photo_path)) 
	Set objFiles = objFolder.Files
	dim curFile
	For each curFile in objFiles
		objFS.DeleteFile(curFile)
	Next	
	' Delete all subfolders and files
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