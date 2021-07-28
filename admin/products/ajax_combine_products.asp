<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
old_id = request("old_productid")
new_id = request("new_product_id")


' Adding the color/info to the details
' Adding main photo into the bottom, and then assigning the details to them


' BEGIN MODIFYING NEW PRODUCT ========================
	' Add comments to new product
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET ProductNotes = (CAST(ProductNotes AS NVARCHAR(MAX)) + CAST('' + NCHAR(13) + ' Combined FROM product #" & old_id & " on " & date() & "' AS NVARCHAR(MAX))) WHERE productID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("new_id",3,1,10,new_id))
	objCmd.Execute()


' BEGIN MODIFYING OLD PRODUCT ========================

	' Retrieve values for images in old_id product to auto assign the newly moved details to a photo 
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "select largepic, picture from jewelry where productID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("old_id",3,1,10,old_id))
	set rsGetProduct = objCmd.Execute()
	
	' Insert old_id product images into images table, but associated with the new product ID
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_images(product_id, img_full, img_thumb) VALUES (?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("new_id",3,1,15,new_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("img_name",200,1,100,rsGetProduct.Fields.Item("largepic").Value))
	objCmd.Parameters.Append(objCmd.CreateParameter("img_thumb",200,1,100,rsGetProduct.Fields.Item("picture").Value))
	objCmd.Execute()	

	' Get newest image ID for new_id product
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP(1) img_id FROM tbl_images WHERE product_id = ? ORDER BY img_id DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("new_id",3,1,15,new_id))
	set rsGetNewImageID = objCmd.Execute()
		
		img_id = rsGetNewImageID.Fields.Item("img_id").Value
	
	' Update active status and comments on old_id product
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET active = 0, ProductNotes = (CAST(ProductNotes AS NVARCHAR(MAX)) + CAST('' + NCHAR(13) + ' Combined product into product #" & new_id & " on " & date() & "' AS NVARCHAR(MAX))) WHERE productID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("old_id",3,1,10,old_id))
	objCmd.Execute()
	

	' Set image to each detail, and pre-pend detail info from combine form on product_edit page
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET img_id = ?, ProductDetail1 = (? + ' ' + ProductDetail1), ProductID = ? WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("img_id",3,1,10,img_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("detailinfo",200,1,100,request("detailinfo")))
	objCmd.Parameters.Append(objCmd.CreateParameter("new_id",3,1,10,new_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("old_id",3,1,10,old_id))
	objCmd.Execute()
	
	'Move any extra stock images that didn't get moved into the new listing
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE tbl_images SET  product_id = ? WHERE product_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("new_id",3,1,10,new_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("old_id",3,1,10,old_id))
	objCmd.Execute()
		
	' Move reviews & photos from old product to new product
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_PhotoGallery SET orig_productid = ProductID, orig_detailid = DetailID, ProductID = ? WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("new_id",3,1,10,new_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("old_id",3,1,10,old_id))
	objCmd.Execute()
	

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBLReviews SET orig_productid = ProductID, orig_detailid = ISNULL(DetailID,0), ProductID = ? WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("new_id",3,1,10,new_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("old_id",3,1,10,old_id))
	objCmd.Execute()
	

DataConn.Close()
%>
