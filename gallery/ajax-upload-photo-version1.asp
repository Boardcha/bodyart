<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
Set Upload = Server.CreateObject("Persits.Upload")

	' 	Live server path
	'	photo_path = "C:\inetpub\wwwroot\gallery\uploads"
		photo_path = "F:\GalleryUploads"

	'	Localhost Path
	'	photo_path = "C:\inetpub\wwwroot\bootstrap-svn\gallery\uploads"	
	'	photo_path = "C:\inetpub\wwwroot\BAF_Site_ASP\gallery\uploads"	
	
	' Upload photo
	Upload.OverwriteFiles = False
	Upload.Save(photo_path)
	
	
	' Check for photo being submitted TWICE on the same order detail ID # coming from the account history page
	If Upload.Form("order-detailid") <> "" then
		ItemNumber = Upload.Form("order-detailid")
	Else
		ItemNumber = 1
	end if


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT OrderDetailID, photoID FROM TBL_PhotoGallery WHERE OrderDetailID = ? AND OrderDetailID <> 0"
	objCmd.Parameters.Append(objCmd.CreateParameter("@OrderDetailID",3,1,10,ItemNumber))
	Set rsCheckDuplicate = objCmd.Execute()

	' Duplicate order detail #, not file name
	If rsCheckDuplicate.EOF Then 
	

	' Retrieve newest filename (in case of a duplicate)
	For Each File in Upload.Files
		filename = File.Filename
	next
	
	' Resize original image to MAX width 1,000 pixels
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Jpeg.PreserveMetadata = True ' makes sure mobile device images are facing the correct way
	Jpeg.Open photo_path & "\" & filename
	Jpeg.ApplyOrientation ' makes sure mobile device images are facing the correct way
	' New width
	L = 1000
	' Resize preserve aspect ratio
	Jpeg.Width = L
	Jpeg.Height = Jpeg.OriginalHeight * L / Jpeg.OriginalWidth
	Jpeg.Save photo_path & "\" & filename
	set Jpeg = nothing
	
	' Create and save thumbnail
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Jpeg.PreserveMetadata = True ' makes sure mobile device images are facing the correct way
	Jpeg.Open photo_path & "\" & filename
	Jpeg.ApplyOrientation ' makes sure mobile device images are facing the correct way
	' New width
	L = 300
	' Resize, preserve aspect ratio
	Jpeg.Width = L
	Jpeg.Height = Jpeg.OriginalHeight * L / Jpeg.OriginalWidth
	Jpeg.Save photo_path & "\thumb_" & filename
	set Jpeg = nothing
	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_PhotoGallery (filename, ProductID, customerID, DetailID, OrderDetailID, Name, Email, DateSubmitted) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("filename",200,1,250,filename))
	objCmd.Parameters.Append(objCmd.CreateParameter("@ProductID",3,1,10,Upload.Form("productid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("@CustomerID",3,1,10,Upload.Form("custid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("@DetailID",3,1,15,Upload.Form("photo-detailid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("@OrderDetailID",3,1,15,Upload.Form("order-detailid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("@Name",200,1,50,Upload.Form("photo-name")))
	objCmd.Parameters.Append(objCmd.CreateParameter("@Email",200,1,75,Upload.Form("photo-email")))
	objCmd.Parameters.Append(objCmd.CreateParameter("@DateSubmitted",135,1,30, now() ))
	objCmd.Execute()

	if Upload.Form("order-detailid") <> 1 then
		' set status photographed to Y on order detail ID
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_OrderSummary SET ProductPhotographed = 'Y' WHERE OrderDetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("photo_id",3,1,15, Upload.Form("order-detailid")))
		objCmd.Execute()
	end if

%>
	{
		"status":"success",
		"status_text":"Photo successfully uploaded. Please give us a few days until it's reviewed for approval.",
		"order_id":"<%= ItemNumber %>"
	}
<%
			

else ' if a duplicate has been detected
%>
	{
		"status":"duplicate",
		"status_text":"A photo has already been submitted for this item on your order.",
		"order_id":"<%= ItemNumber %>"
	}
<%
End If 	' rsCheckDuplicate.EOF  Only checks for duplicate of a photo submitted for a particular item from the order history page

Set Upload = nothing
DataConn.Close()
%>
