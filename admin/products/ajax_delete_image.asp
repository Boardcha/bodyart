<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../../Connections/chilkat.asp" -->
<!--#include file="../../Connections/aws-s3.asp" -->
<%

'===== GET IMAGE FileName
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM tbl_images WHERE img_id = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,Request.QueryString("imgid")))
Set rsImages = objCmd.Execute()

If NOT rsImages.eof then
	db_full_filename = rsImages.Fields.Item("img_full").Value
	db_thumb_filename = rsImages.Fields.Item("img_thumb").Value

	' This example assumes the Chilkat HTTP API to have been previously unlocked
	' See Global Unlock Sample for sample code

	set http = Server.CreateObject("Chilkat_9_5_0.Http")
	' Insert your access key here
	http.AwsAccessKey = AWS_ACCESSKEY
	' Insert your secret key here
	http.AwsSecretKey = AWS_SECRETKEY
	
	http.KeepResponseBody = 1
	
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM jewelry WHERE largepic = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("largepic",200,1,200,db_full_filename))
	Set rsImage = objCmd.Execute()	
	'Check if the image does not exist in DB as a main image
	If Not rsImage.EOF Then
		doesImageNameExistinDB = true
	End If
		
	If rsImages("is_video") = 1 Then
		success1 = http.S3_DeleteObject("baf-videos", db_full_filename)
	Else	
		If doesImageNameExistinDB = false Then
			success1 = http.S3_DeleteObject("bodyartforms-products", db_full_filename)
		End If
	End If 
	
	If doesImageNameExistinDB = false Then
		success2 = http.S3_DeleteObject("bodyartforms-products", db_thumb_filename)
		success3 = http.S3_DeleteObject("baf-thumbs-400", db_thumb_filename)
	End If
	
	If (success1 <> 1 OR success2 <> 1 OR success3 <> 1) Then
		'Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
		'Response.End
	End If

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM tbl_images WHERE img_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("imgid",3,1,15,request.queryString("imgid")))
	objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET img_id = 0  WHERE img_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("imgid",3,1,15,request.queryString("imgid")))
	objCmd.Execute()	
end if	
DataConn.Close()
%>