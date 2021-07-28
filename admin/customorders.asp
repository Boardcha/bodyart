<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="UCII_Cart.asp"-->
<%
' UltraDev Shopping Cart II
' Copyright (c) 2001 Joseph Scavitto All Rights Reserved
' www.thechocolatestore.com/ultradev
Dim UCII_CartColNames,UCII_ComputedCols,UCII__i
UCII_CartColNames = Array("Product_Details","ProductDetailsID","ProductID","Quantity","Name","Price","Total")
UCII_ComputedCols = Array("","","","","","","Price")
Set UCII = VBConstuctCart("UCartII",0,UCII_CartColNames,UCII_ComputedCols)
UCII__i = 0
%>
<%
Dim qryAddtoCart__MMColParam
qryAddtoCart__MMColParam = "0"
if (Request.form("prod_details")    <> "") then qryAddtoCart__MMColParam = Request.form("prod_details")   
%>
<%
Dim qryAddtoCart__MMColParam2
qryAddtoCart__MMColParam2 = "0"
if (Request.form("UCII_recordId")      <> "") then qryAddtoCart__MMColParam2 = Request.form("UCII_recordId")     
%>
<%
set qryAddtoCart = Server.CreateObject("ADODB.Recordset")
qryAddtoCart.ActiveConnection = MM_bodyartforms_sql_STRING
qryAddtoCart.Source = "SELECT *  FROM ProductDetails  WHERE ProductDetailID=" + Replace(qryAddtoCart__MMColParam, "'", "''") + " AND ProductID=" + Replace(qryAddtoCart__MMColParam2, "'", "''") + ";"
qryAddtoCart.CursorLocation = 3 'adUseClient
qryAddtoCart.LockType = 1 'Read-only records
qryAddtoCart.Open()
qryAddtoCart_numRows = 0
%>
<%
Dim rsGetRecords
Dim rsGetRecords_numRows

Set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT * FROM jewelry"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()

rsGetRecords_numRows = 0
%>
<%
' UltraCart II Add To Cart Via Form Version 1.00  
UCII_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
	UCII_editAction = UCII_editAction & "?" & Request.QueryString
End If
UCII_recordId = CStr(Request.Form("UCII_recordId"))
If (Request.Form("UCII_recordId").Count = 1) Then
	set UCII_rs=rsGetRecords
	UCII_uniqueCol="ProductID"
	If (NOT (UCII_rs is Nothing)) Then
		If (UCII_rs.Fields.Item(UCII_uniqueCol).Value <> UCII_recordId) Then
			If (UCII_rs.CursorType > 0) Then
				If (Not UCII_rs.BOF) Then UCII_rs.MoveFirst
			Else
				UCII_rs.Close
				UCII_rs.Open
			End If
		Do While (Not UCII_rs.EOF)
			If (Cstr(UCII_rs.Fields.Item(UCII_uniqueCol).Value) = UCII_recordId) Then
				Exit Do
			End If
			UCII_rs.MoveNext
		Loop
		End If
	End If
	UCII.AddItem UCII_rs,Array("FORM","NONE","NONE","FORM","FORM","FORM","NONE"),Array("detail","0","ProductID","qty","desc","price",""),"increment"
	UCII_redirectPage = "../viewcart.asp"
	If (UCII_redirectPage <> "") Then
		If (InStr(1, UCII_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
			UCII_redirectPage = UCII_redirectPage & "?" & Request.QueryString
		End If
	Call Response.Redirect(UCII_redirectPage)
	End If
End If
%>
<html>
<title>Add item to cart</title>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
</head>
<body bgcolor="#666699" topmargin="0" text="#CCCCCC" link="#CCCCCC" vlink="#CCCCCC">
<!--#include file="admin_header.asp"-->
<span class="adminheader">Add custom item to order</span> 
<form action="<%=UCII_editAction%>" name="form1" method="post">
  <p> 
    <input type="hidden" name="detail" value="-">
    
    <font face="Arial" size="2"> <font face="Verdana"> <font size="1"> 
    <input type="text" name="qty" size="1" value="1">
    &nbsp;&nbsp;&nbsp;Desc 
    <input type="text" name="desc" value="" size="30">
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$ 
    <input type="text" name="price" size="5">
    </font></font></font> <font face="Verdana" size="1"> 
    <input class="detailsbutton" type="submit" name="Submit" value="ADD TO CART">
    </font> </p>
  <input type="hidden" name="UCII_recordId" value="1">
</form>
</body>
</html>
<%
qryAddtoCart.Close()
%>
<%
rsGetRecords.Close()
Set rsGetRecords = Nothing
%>
