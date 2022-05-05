<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../Connections/chilkat.asp" -->
<!--#include file="../../Connections/aws-s3.asp" -->
<!--#include file="inc_content_type.asp" -->
<%
	Set Upload = Server.CreateObject("Persits.Upload")
	'Upload.IgnoreNoPost = True
	photo_path = "img_temp"
	Upload.OverwriteFiles = True
	Upload.SaveVirtual(photo_path)

	i=0
	While i < 3
	
		Set File = Upload.Files("file[" & i & "]")	
		If Not File Is Nothing Then
			File_ImgWidth = getImageWidth("img_temp\" & File.fileName)

			'If there is already images in DB, use the file name in DB
			If Upload.form("selected_img_id") <> "" Then
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "SELECT product_id, img_full, img_thumb, img_thumb_400, img_description FROM tbl_images WHERE img_id = " &  Upload.form("selected_img_id")
				Set rsImages = objCmd.Execute()
				If Not rsImages.EOF Then
					If File_ImgWidth <= 95 Then
						fileName = rsImages("img_thumb")
					ElseIf File_ImgWidth <= 405 Then	
						fileName = rsImages("img_thumb_400")
					Else ' 1000x1000	
						fileName = rsImages("img_full")
					End If
				Else
					Call ReleaseConnection()
					Response.Write "Error"
					Response.End
				End If
			Else
				'If there are no images in DB, generate a file name
				If File_ImgWidth <= 95 Then
					fileName = "thumbnail-90-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(File.FileName, InstrRev(File.FileName, ".") + 1)				
					fileName_90 = fileName
				ElseIf File_ImgWidth <= 405 Then	
					fileName = "thumbnail-400-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(File.FileName, InstrRev(File.FileName, ".") + 1)
					fileName_400 = fileName
				Else '1000x1000
					fileName = File.FileName
					fileName_1000 = fileName
				End If	
			End If 	

			If File_ImgWidth <= 95 Then '90x90
				File.Copy Server.MapPath(photo_path & "\90x90\" & fileName)
				ImageLocation = photo_path & "\90x90\" & fileName
			elseif File_ImgWidth <= 405 Then '400x400
				File.Copy Server.MapPath(photo_path & "\400x400\" & fileName)
				ImageLocation = photo_path & "\400x400\" & fileName
			else '1000x1000
				File.Copy Server.MapPath(photo_path & "\1000x1000\" & fileName)
				ImageLocation = photo_path & "\1000x1000\" & fileName
			end if
			File.Delete	
			
			
			' Upload files to S3 bucket
			set objFs=Server.CreateObject("Scripting.FileSystemObject")
			If objFs.FileExists(Server.MapPath(ImageLocation)) Then
				set http = Server.CreateObject("Chilkat_9_5_0.Http")
				' Insert your AWS keys here:
				http.AwsAccessKey = AWS_ACCESSKEY
				http.AwsSecretKey = AWS_SECRETKEY
				contentType = getContentTypeFromFileName(fileName) ' This is mandatory otherwise S3 cannot register mime type		
				If Instr(ImageLocation, "\400x400\") > 0 Then bucketName = "baf-thumbs-400" Else bucketName = "bodyartforms-products"		
				success = http.S3_UploadFile(Replace(Server.MapPath(ImageLocation), "\", "/"), contentType, bucketName, fileName)
			End If
			
			If success = 1 Then
				If Instr(ImageLocation, "\90x90\") > 0 Then
					successfullyUploadedImages = successfullyUploadedImages & "90x90" & ", "
				ElseIf Instr(ImageLocation, "\400x400\") > 0 Then
					successfullyUploadedImages = successfullyUploadedImages & "400x400" & ", "
				ElseIf Instr(ImageLocation, "\1000x1000\") > 0 Then
					successfullyUploadedImages = successfullyUploadedImages & "1000x1000" & ", "
				End If
			End If	
		End If		
	
		i = i + 1	
	Wend
	
			
	If fileName_90 <> "" And fileName_400 <> "" And fileName_1000 <> "" Then
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_images(product_id, img_full, img_thumb_400, img_thumb, img_description) VALUES (?, ?, ?, ?, ?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("productid" ,3 ,1, 15, Upload.form("productid")))
		objCmd.Parameters.Append(objCmd.CreateParameter("img_full", 200, 1, 100, fileName_1000))
		objCmd.Parameters.Append(objCmd.CreateParameter("img_thumb_400", 200, 1, 100, fileName_400))
		objCmd.Parameters.Append(objCmd.CreateParameter("img_thumb", 200, 1, 100, fileName_90))
		objCmd.Parameters.Append(objCmd.CreateParameter("img_description", 200, 1, 50, Upload.form("add_img_description")))
		objCmd.Execute()
	End If
	
	'Remove last comma from successfullyUploadedImages
	If Len(successfullyUploadedImages) > 0 Then successfullyUploadedImages = Left(successfullyUploadedImages, Len(successfullyUploadedImages) - 2)
	Response.Write "{""uploadedimages"" : """ & successfullyUploadedImages & """}"
	
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

Call ReleaseConnection()

Function ReleaseConnection
	set objFs = nothing	
	Set Upload = nothing
	DataConn.Close()
End Function
Function getImageWidth(file)
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Path = Server.MapPath(file)
	Jpeg.Open Path	
	getImageWidth = Jpeg.OriginalWidth
	Set Jpeg = Nothing
End Function

%>
