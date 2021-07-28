<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../Connections/chilkat.asp" -->
<!--#include file="../../Connections/aws-s3.asp" -->
<%

sliderId = Session("sliderId")

'===== GET IMAGE FileName
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Sliders WHERE sliderID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("sliderID",3,1,10,sliderId))
Set rs_getImages = objCmd.Execute()

set http = Server.CreateObject("Chilkat_9_5_0.Http")
' Insert your access key here
http.AwsAccessKey = AWS_ACCESSKEY
' Insert your secret key here
http.AwsSecretKey = AWS_SECRETKEY
bucketName = "baf-hero-images"
http.KeepResponseBody = 1

If NOT rs_getImages.EOF Then
	If Not IsNull(rs_getImages.Fields.Item("img550x350").Value) And rs_getImages.Fields.Item("img550x350").Value<>"" Then succes = http.S3_DeleteObject(bucketName, rs_getImages.Fields.Item("img550x350").Value)
	If Not IsNull(rs_getImages.Fields.Item("img850x350").Value) And rs_getImages.Fields.Item("img850x350").Value<>"" Then succes = http.S3_DeleteObject(bucketName, rs_getImages.Fields.Item("img850x350").Value)
	If Not IsNull(rs_getImages.Fields.Item("img1024x350").Value) And rs_getImages.Fields.Item("img1024x350").Value<>"" Then succes = http.S3_DeleteObject(bucketName, rs_getImages.Fields.Item("img1024x350").Value)
	If Not IsNull(rs_getImages.Fields.Item("img1600x350").Value) And rs_getImages.Fields.Item("img1600x350").Value<>"" Then succes = http.S3_DeleteObject(bucketName, rs_getImages.Fields.Item("img1600x350").Value)
	If Not IsNull(rs_getImages.Fields.Item("img1920x350").Value) And rs_getImages.Fields.Item("img1920x350").Value<>"" Then succes = http.S3_DeleteObject(bucketName, rs_getImages.Fields.Item("img1920x350").Value)
End If
	
DataConn.Close()
%>