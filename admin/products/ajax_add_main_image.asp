<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../Connections/chilkat.asp" -->
<!--#include file="../../Connections/aws-s3.asp" -->
<!--#include file="inc_content_type.asp" -->
<!--#include file="../../functions/random_integer.asp" -->
<%
	Set Upload = Server.CreateObject("Persits.Upload")
	photo_path = "img_temp"
	Upload.OverwriteFiles = True
	Upload.SaveVirtual(photo_path)
	
	i=0
	While i < 3
	
		Set File = Upload.Files("file[" & i & "]")	
		If Not File Is Nothing Then
			File_ImgWidth = getImageWidth("img_temp\" & File.fileName)

			'If there is already images in DB, use the file name in DB
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "SELECT picture, picture_400, largepic FROM jewelry WHERE ProductID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("ID", 3, 1, 10, Upload.form("productid")))
			Set rsImages = objCmd.Execute()
			If Not rsImages.EOF Then
				If File_ImgWidth <= 95 Then
					fileName = rsImages("picture")
				ElseIf File_ImgWidth <= 405 Then	
					fileName = rsImages("picture_400")
				Else ' 1000x1000	
					fileName = rsImages("largepic")
				End If
			Else
				Call ReleaseConnection()
				Response.End
			End If
			
			'If there are no images in DB, generate a file name
			If  fileName = "nopic.gif" Then
				If File_ImgWidth < 95 Then
					fileName = "thumbnail-90-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(File.FileName, InstrRev(File.FileName, ".") + 1)				
				ElseIf File_ImgWidth < 405 Then	
					fileName = "thumbnail-400-" & Month(date) & Day(date) & Year(date) & Hour(time) & Minute(time) & Second(time) & "-" & Upload.form("productid") & "." & Mid(File.FileName, InstrRev(File.FileName, ".") + 1)
				Else '1000x1000
					fileName = File.FileName
				End If	
			End If 	
			
			If File_ImgWidth < 95 Then '90x90
				File.Copy Server.MapPath(photo_path & "\90x90\" & fileName)
				ImageLocation = photo_path & "\90x90\" & fileName
			elseif File_ImgWidth < 405 Then '400x400
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
					field = "picture"
					' If it is 90x90, return image path for updating thumbnail on product-edit.asp
					thumbnail = """thumbnail"" : ""https://s3.amazonaws.com/bodyartforms-products/" & fileName & "?ver=" & getInteger(8) & """, "
				ElseIf Instr(ImageLocation, "\400x400\") > 0 Then
					successfullyUploadedImages = successfullyUploadedImages & "400x400" & ", "
					field = "picture_400"
				ElseIf Instr(ImageLocation, "\1000x1000\") > 0 Then
					successfullyUploadedImages = successfullyUploadedImages & "1000x1000" & ", "
					field = "largepic"
				End If
			
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE jewelry SET " & field & " = ? WHERE productID=?"
				objCmd.Parameters.Append(objCmd.CreateParameter("fileName", 200, 1, 100, fileName))
				objCmd.Parameters.Append(objCmd.CreateParameter("productID", 3, 1 ,15, Upload.form("productid")))
				objCmd.Execute()
			End If	
		End If
		i = i + 1	
	Wend
	
	'Remove last comma from successfullyUploadedImages
	If Len(successfullyUploadedImages) > 0 Then successfullyUploadedImages = Left(successfullyUploadedImages, Len(successfullyUploadedImages) - 2)
	Response.Write "{" & thumbnail & """uploadedimages"" : """ & successfullyUploadedImages & """}"
	
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
