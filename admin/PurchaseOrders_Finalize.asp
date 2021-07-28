<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
Dim rsGetRestockItems__MMColParam
rsGetRestockItems__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsGetRestockItems__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsGetRestockItems
Dim rsGetRestockItems_cmd
Dim rsGetRestockItems_numRows

Set rsGetRestockItems_cmd = Server.CreateObject ("ADODB.Command")
rsGetRestockItems_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRestockItems_cmd.CommandText = "SELECT * FROM dbo.QRY_PORestock WHERE PurchaseOrderID = ? ORDER BY title ASC" 
rsGetRestockItems_cmd.Prepared = true
rsGetRestockItems_cmd.Parameters.Append rsGetRestockItems_cmd.CreateParameter("param1", 5, 1, -1, rsGetRestockItems__MMColParam) ' adDouble

Set rsGetRestockItems = rsGetRestockItems_cmd.Execute
rsGetRestockItems_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetRestockItems_numRows = rsGetRestockItems_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsGetRestockItems_total
Dim rsGetRestockItems_first
Dim rsGetRestockItems_last

' set the record count
rsGetRestockItems_total = rsGetRestockItems.RecordCount

' set the number of rows displayed on this page
If (rsGetRestockItems_numRows < 0) Then
  rsGetRestockItems_numRows = rsGetRestockItems_total
Elseif (rsGetRestockItems_numRows = 0) Then
  rsGetRestockItems_numRows = 1
End If

' set the first and last displayed record
rsGetRestockItems_first = 1
rsGetRestockItems_last  = rsGetRestockItems_first + rsGetRestockItems_numRows - 1

' if we have the correct record count, check the other stats
If (rsGetRestockItems_total <> -1) Then
  If (rsGetRestockItems_first > rsGetRestockItems_total) Then
    rsGetRestockItems_first = rsGetRestockItems_total
  End If
  If (rsGetRestockItems_last > rsGetRestockItems_total) Then
    rsGetRestockItems_last = rsGetRestockItems_total
  End If
  If (rsGetRestockItems_numRows > rsGetRestockItems_total) Then
    rsGetRestockItems_numRows = rsGetRestockItems_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsGetRestockItems_total = -1) Then

  ' count the total records by iterating through the recordset
  rsGetRestockItems_total=0
  While (Not rsGetRestockItems.EOF)
    rsGetRestockItems_total = rsGetRestockItems_total + 1
    rsGetRestockItems.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsGetRestockItems.CursorType > 0) Then
    rsGetRestockItems.MoveFirst
  Else
    rsGetRestockItems.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsGetRestockItems_numRows < 0 Or rsGetRestockItems_numRows > rsGetRestockItems_total) Then
    rsGetRestockItems_numRows = rsGetRestockItems_total
  End If

  ' set the first and last displayed record
  rsGetRestockItems_first = 1
  rsGetRestockItems_last = rsGetRestockItems_first + rsGetRestockItems_numRows - 1
  
  If (rsGetRestockItems_first > rsGetRestockItems_total) Then
    rsGetRestockItems_first = rsGetRestockItems_total
  End If
  If (rsGetRestockItems_last > rsGetRestockItems_total) Then
    rsGetRestockItems_last = rsGetRestockItems_total
  End If

End If
%>
<%
' *** FX Update Multiple Records in FRM_UpdateStock
If (rsGetRestockItems_first <> "") Then upd_rsGetRestockItems = rsGetRestockItems_first-1 Else upd_rsGetRestockItems = 0 End If ' counter
If (cStr(Request.Form("button")) <> "") Then
  FX_sqlerror = ""
  FX_updredir = "PurchaseOrders.asp"
  tmp = "ADODB.Command"
  Set update_Multi = Server.CreateObject(tmp)
  update_Multi.ActiveConnection = MM_bodyartforms_sql_STRING

  For N = upd_rsGetRestockItems+1 To rsGetRestockItems_total
      If (Request.Form("QtyAdd"&N) <> "") Then s1 = Replace(Request.Form("QtyAdd"&N),"'","''") Else s1 = "0" End If
	  If (Request.Form("QtyOrig"&N) <> "") Then s2 = Replace(Request.Form("QtyOrig"&N),"'","''") Else s2 = "0" End If
      If (Request.Form("fx_updmatch"&N) <> "") Then sw = Replace(Request.Form("fx_updmatch"&N),"'","''") Else sw = "0" End If
	  If (s2 = 0 AND s1 > 0) then ReStock = ", DateRestocked = '"& date() &"'" else ReStock = "" End if
	   
    On Error Resume Next
      update_Multi.CommandText = "UPDATE dbo.ProductDetails SET qty="+s1+" + "+s2+" "+Restock+" WHERE ProductDetailID="+sw+""
      update_Multi.Execute

    If (Err.Description <> "") Then
      FX_sqlerror = FX_sqlerror & "Row " & N & ": " & Err.Description & "<br><br>"
    End If
  Next 		
  update_Multi.ActiveConnection.Close
  
	' update order to be received status
set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "UPDATE dbo.TBL_PurchaseOrders SET Received='Y', DateReceived='"& date() &"' WHERE PurchaseOrderID="+ Request.Querystring("ID")+"" 
commUpdate.Execute()
  
  thispath = cStr(Request.ServerVariables("SCRIPT_NAME"))
  If (FX_updredir = "") Then FX_updredir = Mid(thispath, InstrRev(thispath, "/")+1) End If
  If (Request.QueryString <> "") Then
    ch = "&"
    If (InStr(FX_updredir,"?") = 0) Then ch = "?" End If	
    FX_updredir = FX_updredir & ch & Request.QueryString
  End If
  If (FX_sqlerror <> "") Then
    Response.Write("<font color=""red"">"&FX_sqlerror&"</font>")
  Else Response.Redirect(FX_updredir) End If
End If
%>
<html>
<head>

<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Put items in stock</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body topmargin="0" class="MainBkgdColor">
<!--#include file="admin_header.asp"-->
<span class="adminheader">Put items in stock<br>
</span>&nbsp;&nbsp;
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td colspan="2"> <div align="left" class="faqs">
      <% If Not rsGetRestockItems.EOF Or Not rsGetRestockItems.BOF Then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_UpdateStock" id="FRM_UpdateStock">          
  <table width="60%" border="0" cellspacing="1" cellpadding="6">
            <tr valign="middle" bgcolor="#000000">
              <td align="right" bgcolor="#000000" class="faqs"><input type="submit" name="button" id="button" value="PUT ITEMS IN STOCK"></td>
            </tr>
              <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetRestockItems.EOF)) 
%>
                <tr valign="middle" bgcolor="#ececec">
                  <td bgcolor="#ececec"><p>
                    <% upd_rsGetRestockItems = upd_rsGetRestockItems+1 %>
                      <input type="hidden" name="fx_updmatch<%=upd_rsGetRestockItems%>" value="<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%>">
                      <span class="materialText"><%=(rsGetRestockItems.Fields.Item("POAmount").Value)%></span>
                      &nbsp;
                      <input name="QtyAdd<%=upd_rsGetRestockItems%>" type="hidden" id="QtyAdd" value=<%=(rsGetRestockItems.Fields.Item("POAmount").Value)%>  >
                    <% if (rsGetRestockItems.Fields.Item("qty").Value) < 0 then %>
                    <input type="hidden" name="QtyOrig<%=upd_rsGetRestockItems%>" value=0>
                    <% else %>
                    <input type="hidden" name="QtyOrig<%=upd_rsGetRestockItems%>" value=<%=(rsGetRestockItems.Fields.Item("qty").Value)%>>
                    <% end if %>
&nbsp; <a href="product-edit.asp?ProductID=<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>&info=less" target="_blank" class="EditSelect_Links"><%=(rsGetRestockItems.Fields.Item("title").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("gauge").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("ProductDetail1").Value)%></a>
<input name="OrderID" type="hidden" id="OrderID" value="<%= Request.Querystring("ID") %>">
                  </p>                  </td>
              </tr>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetRestockItems.MoveNext()
Wend
%>
          </table>
  </form>
          <% End If ' end Not rsGetRestockItems.EOF Or NOT rsGetRestockItems.BOF %>
      <% If rsGetRestockItems.EOF And rsGetRestockItems.BOF Then %>
        <p class="adminheader"><font color="#FFFF00">No order to display</font></p>
        <% End If ' end rsGetRestockItems.EOF And rsGetRestockItems.BOF %>
</div></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
rsGetRestockItems.Close()
Set rsGetRestockItems = Nothing
%>
