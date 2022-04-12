<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	'===== ONLY REASSIGN ITEMS THAT ARE IN THE GOLD CASE OR HAVE A BIN # OF 0. THIS WILL PREVENT ITEMS THAT ARE ALREADY IN LIMITED BINS FROM BEING MOVED.=============================

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
	objCmd.CommandText = "UPDATE ProductDetails SET BinNumber_Detail = CASE WHEN BinNumber_Detail = 37 OR BinNumber_Detail = 0 THEN ? ELSE BinNumber_Detail END, active = 1, DetailCode = 0 WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("bin_number",3,1,15,request.form("bin")))
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15,request.form("productid")))
	objCmd.Execute()


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = MM_bodyartforms_sql_STRING  
	objCmd.CommandText = "UPDATE jewelry SET active = 1, pull_completed = 1, date_pulled = GETDATE() FROM jewelry WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request.form("productid")))
	objCmd.Execute()

	' Pull all details with product to update each edits log
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = MM_bodyartforms_sql_STRING  
	objCmd.CommandText = "SELECT * FROM ProductDetails WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request.form("productid")))
	set rsGetProductDetails =  objCmd.Execute()

	while NOT rsGetProductDetails.EOF

	'Write info to edits log	
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, product_id, detail_id, description, edit_date) VALUES (?," & request.form("productid") & "," & rsGetProductDetails("ProductDetailID") & ",'Automated message - Moved item using pulling discontinued app. Moved to BIN " & request.form("bin") & "','" & now() & "')"
	objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser("user_id") ))
	objCmd.Execute()

	rsGetProductDetails.movenext()
	wend


DataConn.Close()
Set objCmd = Nothing
%>