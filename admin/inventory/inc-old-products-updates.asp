<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
var_column = request.form("column")
var_value = request.form("value")
var_productid = request.form("productid")


	set rsOriginalRecord = Server.CreateObject("ADODB.Recordset")
	rsOriginalRecord.ActiveConnection = MM_bodyartforms_sql_STRING
	rsOriginalRecord.Source = "SELECT SaleDiscount, ProductNotes, type FROM jewelry WHERE ProductID = " & var_productid
	rsOriginalRecord.CursorLocation = 3 'adUseClient
	rsOriginalRecord.LockType = 1 'Read-only records
	rsOriginalRecord.Open()

    '===== IF TYPE FIELD UPDATED THEN UPDATE TYPE FIELD IN JEWELRY.TYPE & WRITE NOTES =============
    if var_column = "type" then

        Notes = rsOriginalRecord.Fields.Item("ProductNotes") & "<br>--- Updated status from " & rsOriginalRecord.Fields.Item("type") & " to " & var_value & " by " & user_name & " on " & Now()

        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE jewelry SET type = ?, ProductNotes = ? WHERE ProductID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("value",8,1,30,  var_value))
        objCmd.Parameters.Append(objCmd.CreateParameter("notes",8,1,8000,  Notes))
        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,15, var_productid ))
        objCmd.Execute()
        
    end if

    '===== IF SALEDISCOUNT FIELD UPDATED THEN UPDATE TYPE FIELD IN JEWELRY.SaleDiscount & WRITE NOTES =============
	if var_column = "saleDiscount" then

        Notes = rsOriginalRecord.Fields.Item("ProductNotes") & "<br>--- Updated sale amount from " & rsOriginalRecord.Fields.Item("SaleDiscount") & " to " & var_value & "% by " & user_name & " on " & Now()

        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE jewelry SET SaleDiscount = ?, ProductNotes = ? WHERE ProductID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("discount",6,1,10,  var_value))
        objCmd.Parameters.Append(objCmd.CreateParameter("notes",8,1,8000,  Notes))
        objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, var_productid ))
        objCmd.Execute()

    end if
%>