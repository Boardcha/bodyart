<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
response.write request.form("qty-onhand")

if request.form("active") = 1 then
	active = 1
else
	active = 0
end if

	if request.form("colors_add") <> "" then
		'break values out by comma and then reformat to be full text search friendly before saving into field
		value_array =split(request.form("colors_add"),",")
			For Each strItem In value_array
				if strItem <> "" then 
					var_colors = var_colors & "" & strItem & " "
				end if 			
			Next
	end if
	
	if request.form("materials_add") <> "" then
		'break values out by comma and then reformat to be full text search friendly before saving into field
		value_array =split(request.form("materials_add"),",")
			For Each strItem In value_array
				if strItem <> "" then 
					var_materials = var_materials & " , " & strItem & ""
				end if 			
			Next
	end if

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO ProductDetails (item_order, DetailCode, location, qty, stock_qty, restock_threshold, Gauge, Length, ProductDetail1, price, wlsl_price, detail_code, ProductID, active, last_inactivation_date, weight, colors, detail_materials, wearable_material, DateAdded) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'" & now() & "')"
	objCmd.Parameters.Append(objCmd.CreateParameter("item_order",3,1,10, request.form("sort")))
	objCmd.Parameters.Append(objCmd.CreateParameter("DetailCode",3,1,10,request.form("section")))
	objCmd.Parameters.Append(objCmd.CreateParameter("location",3,1,10,request.form("location")))
	objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,request.form("qty-onhand")))
	objCmd.Parameters.Append(objCmd.CreateParameter("stock_qty",3,1,10,request.form("max")))
	objCmd.Parameters.Append(objCmd.CreateParameter("restock_threshold",3,1,10,request.form("thresh")))
	objCmd.Parameters.Append(objCmd.CreateParameter("Gauge",200,1,225,request.form("gauge")))
	objCmd.Parameters.Append(objCmd.CreateParameter("Length",200,1,225,request.form("length")))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetail1",200,1,225,request.form("detail")))
	objCmd.Parameters.Append(objCmd.CreateParameter("price",6,1,10,request.form("retail")))
	objCmd.Parameters.Append(objCmd.CreateParameter("wlsl_price",6,1,10,request.form("wholesale")))
	objCmd.Parameters.Append(objCmd.CreateParameter("detail_code",200,1,225,request.form("sku")))
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request.form("productid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("active",3,1,10,active))
	If active = 0 Then
		objCmd.Parameters.Append(objCmd.CreateParameter("last_inactivation_date",200,1,30,Cstr(now())))
	Else
		objCmd.Parameters.Append(objCmd.CreateParameter("last_inactivation_date",200,1,30, NULL ))
	End If
	objCmd.Parameters.Append(objCmd.CreateParameter("weight",3,1,10,request.form("weight")))
	objCmd.Parameters.Append(objCmd.CreateParameter("colors",200,1,200,var_colors))
	objCmd.Parameters.Append(objCmd.CreateParameter("materials",200,1,200,var_materials))
	objCmd.Parameters.Append(objCmd.CreateParameter("wearable_material",200,1,200,request.form("wearable_add")))
	objCmd.Execute()

if request.form("location") = 0 then ' reset location if it's not filled out
	
	' retrieve newest product detail ID # to copy into location field
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP 1 ProductDetailID FROM ProductDetails ORDER BY ProductDetailID DESC" 
	Set rsGetID = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET location = ? WHERE ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("location",3,1,10, rsGetID.Fields.Item("ProductDetailID").Value))
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10, rsGetID.Fields.Item("ProductDetailID").Value))
	objCmd.Execute()

end if ' reset location if it's not filled out

DataConn.Close()
%>
{  
   "sort":"23",
   "section":"23",
   "location":"23",
   "qty":"23",
   "max":"23",
   "thresh":"23",
   "gauge":"23",
   "length":"23",
   "detail":"23",
   "retail":"23",
   "wholesale":"23",
   "sku":"23",
   "detailid":"23"
}