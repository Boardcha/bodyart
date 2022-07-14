<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	column_title = request.form("column")
	
	' protection to stop page from sql injection
	if Len(column_title) > 30 then
		response.end
	end if
	
	column_value = request.form("value")
	id = request.form("id")
	friendly_name = request.form("friendly_name")
response.write "column_title: " & column_title & "<br/>"
if request.form("detailid") = "" then ' check to see if a detailid is provided and if not, just update the main products table	


	if column_title = "jewelry" or column_title = "tags" then
		value_array =split(column_value,",")
		column_value = ""
			For Each strItem In value_array
				if strItem <> "" then 
				column_value = column_value & strItem & " "
				end if 				
			Next
	end if
	
	if column_title = "piercing_type" then
		'break values out by comma and then reformat to be full text search friendly before saving into field
		value_array =split(column_value,",")
		column_value = ""
			For Each strItem In value_array
				if strItem <> "" then 
					column_value = column_value & "piercing_type:" & strItem & " "
				end if 			
			Next
	end if
	
	if column_title = "material" or column_title = "internal" or column_title = "flare_type" then
		response.write "PRE: " & column_value & "<br/>"
		'break values out by comma and then reformat to be full text search friendly before saving into field
		value_array =split(column_value,",")
		column_value = ""
			For Each strItem In value_array
				if strItem <> "" then 
					column_value = column_value & " , " & strItem & " "
					response.write "DURING: " & column_value & "<br/>"
				end if 			
			Next
	end if
response.write "AFTER: " & column_value & "<br/>"


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "select " & column_title & " from jewelry where productID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	set rsGetColumn = objCmd.Execute()
	
	orig_value = rsGetColumn.Fields.Item(column_title).Value
	edits_description = "Updated MAIN PRODUCT: " & friendly_name & " from " & orig_value & " to " & column_value
	edits_detail_id = 0
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update jewelry set " & column_title & " = ? where productID = ?"
'	objCmd.Parameters.Append(objCmd.CreateParameter("column",3,1,20,column_title))
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,8000,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	
	If column_title = "active" And column_value = "0" Then
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE jewelry set last_inactivation_date = GETDATE() WHERE productID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
		objCmd.Execute()	
	End If
	
else ' update the details

	detailid = request.form("detailid")
	
	if column_title = "colors" then
		'break values out by comma and then reformat to be full text search friendly before saving into field
		value_array =split(column_value,",")
		column_value = ""
			For Each strItem In value_array
				if strItem <> "" then 
					column_value = column_value & "" & strItem & " "
				end if 			
			Next
	end if
	
	if column_title = "detail_materials" then
		'break values out by comma and then reformat to be full text search friendly before saving into field
		value_array =split(column_value,",")
		column_value = ""
			For Each strItem In value_array
				if strItem <> "" then 
					column_value = column_value & " , " & strItem & " "
				end if 			
			Next
	end if
	
	if column_title = "free_item_expiration_date" And column_value = "" then
		column_value = null
	end if
	if column_title = "free_item_start_date" And column_value = "" then
		column_value = null
	end if	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "select " & column_title & ", ProductDetailID from ProductDetails where ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,detailid))
	set rsGetColumn = objCmd.Execute()
	
	orig_value = rsGetColumn.Fields.Item(column_title).Value
	edits_description = "Updated DETAIL: " & friendly_name & " from " & orig_value & " to " & column_value
	edits_detail_id = rsGetColumn.Fields.Item("ProductDetailID").Value

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update ProductDetails set " & column_title & " = ? where ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,1000,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,detailid))
	objCmd.Execute()
	
	If column_title = "active" And column_value = "0" Then
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails set last_inactivation_date = GETDATE() WHERE ProductDetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,detailid))
		objCmd.Execute()	
	End If	

end if

	if column_title <> "description" or column_title <> "ProductNotes" then ' record edit unless it's the product description -- it's too long to store

		'Write ALL info to edits log table
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, product_id, detail_id, description, edit_date) VALUES (" & user_id & "," & id & ", ?, ?,'" & now() & "')"
		objCmd.Parameters.Append(objCmd.CreateParameter("edits_detail_id",3,1,10, edits_detail_id))
		objCmd.Parameters.Append(objCmd.CreateParameter("edits_description",200,1,1000, edits_description ))
		objCmd.Execute()
	
	end if

Set rsGetUser = nothing
set rsGetColumn = nothing
DataConn.Close()
%>