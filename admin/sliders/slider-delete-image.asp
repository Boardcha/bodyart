<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../Connections/chilkat.asp" -->
<!--#include file="../../Connections/aws-s3.asp" -->
<%

imgId = request.Form("img_id")
sliderId = request.Form("slider_id")

'===== GET IMAGE FileName
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT img" & imgId + " FROM TBL_Sliders WHERE sliderID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("sliderID",3,1,10,sliderId))
Set rs_getImage_Filename = objCmd.Execute()

if NOT rs_getImage_Filename.eof then
	filename = rs_getImage_Filename.Fields.Item("img" & imgId).Value
end if
	
' This example assumes the Chilkat HTTP API to have been previously unlocked
' See Global Unlock Sample for sample code

set http = Server.CreateObject("Chilkat_9_5_0.Http")
' Insert your access key here
http.AwsAccessKey = AWS_ACCESSKEY
' Insert your secret key here
http.AwsSecretKey = AWS_SECRETKEY
bucketName = "baf-hero-images"


http.KeepResponseBody = 1
success = http.S3_DeleteObject(bucketName, filename)


set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_Sliders SET img" & imgId & "=NULL, active=0  WHERE sliderID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("sliderID",3,1,10,sliderId))
objCmd.Execute()	
	
DataConn.Close()
%>